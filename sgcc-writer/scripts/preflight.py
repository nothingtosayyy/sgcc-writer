#!/usr/bin/env python3
"""Read-only preflight for Word tasks.

Checks input, structure, fonts, template, tools and output directory.
Never modifies the input or template. Prints JSON with a green/yellow/red level.
"""

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

FONT_CHAINS = {
    "方正小标宋_GBK": {
        "exact": ["方正小标宋_GBK", "FZXBSK"],
        "fallback": [{"name": "宋体", "aliases": ["宋体", "SimSun", "SIMSUN"]}],
    },
    "仿宋_GB2312": {
        "exact": ["仿宋_GB2312", "FangSong_GB2312", "SIMFANG_GB2312"],
        "fallback": [{"name": "仿宋", "aliases": ["仿宋", "FangSong", "SIMFANG"]}],
    },
    "楷体_GB2312": {
        "exact": ["楷体_GB2312", "KaiTi_GB2312", "SIMKAI_GB2312"],
        "fallback": [{"name": "楷体", "aliases": ["楷体", "KaiTi", "SIMKAI"]}],
    },
    "黑体": {
        "exact": ["黑体", "SimHei", "SIMHEI"],
        "fallback": [],
    },
    "宋体": {
        "exact": ["宋体", "SimSun", "SIMSUN"],
        "fallback": [],
    },
}
REQUIRED_STYLES = [
    "GW-Title", "GW-H1", "GW-H2", "GW-H3", "GW-H4", "GW-Body",
    "GW-TableCaption", "GW-TableText", "GW-TableHeader",
    "GW-FigureCaption", "GW-PageNumber",
]


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def _normalize_font_name(value):
    name = value.strip().lower()
    name = re.sub(r"\s*\((?:truetype|opentype|type 1|type1)\)\s*$", "", name)
    return re.sub(r"[\s_\-]+", "", name)


def _registry_font_names():
    try:
        import winreg
    except ImportError:
        return [], []
    names, sources = [], []
    entries = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", "HKLM registry"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", "HKCU registry"),
    ]
    for hive, key_path, label in entries:
        try:
            key = winreg.OpenKey(hive, key_path)
            count = winreg.QueryInfoKey(key)[1]
            for i in range(count):
                value_name, value_data, _ = winreg.EnumValue(key, i)
                if value_name:
                    names.append(value_name)
                if isinstance(value_data, str) and value_data:
                    names.append(Path(value_data).stem)
            winreg.CloseKey(key)
            sources.append(label)
        except OSError:
            continue
    return names, sources


def _font_dirs():
    dirs = [Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"]
    local = os.environ.get("LOCALAPPDATA")
    if local:
        dirs.append(Path(local) / "Microsoft" / "Windows" / "Fonts")
    programdata = os.environ.get("PROGRAMDATA")
    if programdata:
        dirs.append(Path(programdata) / "Kingsoft" / "office6" / "muifont")
    return dirs


def collect_font_catalog():
    names, sources = _registry_font_names()
    for font_dir in _font_dirs():
        if not font_dir.exists():
            continue
        found = 0
        try:
            for path in font_dir.rglob("*"):
                if path.is_file() and path.suffix.lower() in (".ttf", ".ttc", ".otf"):
                    names.extend([path.name, path.stem])
                    found += 1
        except OSError:
            continue
        if found:
            sources.append(str(font_dir))
    normalized = {_normalize_font_name(n) for n in names if n}
    return {"normalized": normalized, "sources": sources, "entries": len(names)}


def classify_fonts(catalog):
    targets, fallbacks, missing = {}, [], []
    for target, chain in FONT_CHAINS.items():
        exact = next((c for c in chain["exact"] if _normalize_font_name(c) in catalog["normalized"]), None)
        if exact:
            targets[target] = {"status": "exact", "actual": target}
            continue
        fallback = next(
            (item for item in chain["fallback"]
             if any(_normalize_font_name(alias) in catalog["normalized"] for alias in item["aliases"])),
            None,
        )
        if fallback:
            targets[target] = {"status": "fallback", "actual": fallback["name"]}
            fallbacks.append("%s→%s" % (target, fallback["name"]))
        else:
            targets[target] = {"status": "missing", "actual": None}
            missing.append(target)
    return targets, fallbacks, missing


def check_fonts():
    catalog = collect_font_catalog()
    targets, fallbacks, missing = classify_fonts(catalog)
    return {
        "sources": catalog["sources"],
        "installed_entries": catalog["entries"],
        "targets": targets,
        "fallbacks": fallbacks,
        "missing": missing,
    }


def check_template(path):
    result = {"path": str(path), "exists": path.exists(), "styles_missing": []}
    if not path.exists():
        result["styles_missing"] = REQUIRED_STYLES[:]
        return result
    try:
        from docx import Document
        d = Document(str(path))
        names = {s.name for s in d.styles}
        result["styles_missing"] = [s for s in REQUIRED_STYLES if s not in names]
    except Exception as exc:
        result["error"] = str(exc)
        result["styles_missing"] = REQUIRED_STYLES[:]
    return result


def check_docx(path):
    from docx import Document
    d = Document(str(path))
    xml = ""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
        names = z.namelist()
    comments = 1 if "word/comments.xml" in names else 0
    revisions = xml.count("<w:ins ") + xml.count("<w:del ")
    complex_counts = {
        "text_boxes": xml.count("<w:txbxContent"),
        "drawings": xml.count("<w:drawing"),
        "pictures": xml.count("<w:pict"),
        "objects": xml.count("<w:object"),
        "merged_cells": xml.count("<w:gridSpan") + xml.count("<w:vMerge"),
    }
    return {
        "paragraphs": len(d.paragraphs),
        "tables": len(d.tables),
        "images": len(d.inline_shapes),
        "comments": comments,
        "revisions": revisions,
        "complex": complex_counts,
    }


def main():
    parser = argparse.ArgumentParser(description="Read-only preflight for Word tasks")
    parser.add_argument("--input", required=True, help="input .docx/.doc/.docm")
    parser.add_argument("--template", default=None, help="template docx")
    parser.add_argument("--output-dir", default=None, help="output directory")
    parser.add_argument("--json", action="store_true", help="print JSON only")
    args = parser.parse_args()

    src = Path(args.input).expanduser().resolve()
    template = Path(args.template).expanduser().resolve() if args.template else (
        Path(__file__).resolve().parent.parent / "assets" / "templates" / "guowang-preset-v1.0.1.docx"
    )
    out_dir = Path(args.output_dir).expanduser().resolve() if args.output_dir else src.parent

    warnings, errors, notes = [], [], []
    ext = src.suffix.lower()
    input_info = {
        "path": str(src), "exists": src.exists(), "extension": ext,
        "size": None, "sha256": None, "locked_heuristic": False,
    }

    if not src.exists():
        errors.append("E06 输入文件不存在")
    else:
        input_info["size"] = src.stat().st_size
        try:
            input_info["sha256"] = sha256(src)
        except Exception:
            errors.append("E06 输入文件不可读")
        lock = src.parent / ("~$" + src.name)
        input_info["locked_heuristic"] = lock.exists()
        if input_info["locked_heuristic"]:
            errors.append("E01 检测到 Word/WPS 锁文件，文件可能正被占用")

    structure = {}
    if src.exists() and ext == ".docx":
        try:
            structure = check_docx(src)
            if structure["complex"]["text_boxes"] or structure["complex"]["objects"]:
                warnings.append("E05 检测到文本框或嵌入对象，复杂结构可能无法完整保留")
            if structure["complex"]["merged_cells"]:
                warnings.append("E05 检测到复杂合并单元格，需目检")
            if structure["revisions"]:
                notes.append("检测到修订痕迹，排版前确认是否保留")
            if structure["comments"]:
                notes.append("检测到批注，排版时需保留锚点")
        except Exception as exc:
            errors.append("E06 .docx 无法解析：%s" % exc)
    elif src.exists() and ext in (".doc", ".docm"):
        warnings.append("E04 %s 需要先转换为 .docx 再排版" % ext)
    elif src.exists():
        errors.append("E06 不支持的输入格式：%s" % ext)

    tools = {
        "python": sys.version.split()[0],
        "python_docx": True,
        "soffice": shutil.which("soffice"),
        "winword": shutil.which("winword"),
        "wps": shutil.which("wps"),
    }
    if not any([tools["soffice"], tools["winword"], tools["wps"]]):
        notes.append("未检测到命令行 Word/WPS/LibreOffice；若宿主提供 Word 能力可忽略")

    fonts = check_fonts()
    if fonts["fallbacks"]:
        notes.append("字体回退：%s；按实际字体排版并写入回执" % "、".join(fonts["fallbacks"]))
    if fonts["missing"]:
        warnings.append("E02 字体缺失且无回退：%s；将按字体名请求，由 Word 替换" % "、".join(fonts["missing"]))

    tmpl = check_template(template)
    if not tmpl["exists"]:
        warnings.append("E07 模板不存在：%s；将按参数手工排版" % template)
    elif tmpl.get("styles_missing"):
        warnings.append("E07 模板样式缺失：%s" % "、".join(tmpl["styles_missing"]))

    output = {"dir": str(out_dir), "writable": os.access(str(out_dir), os.W_OK) if out_dir.exists() else os.access(str(out_dir.parent), os.W_OK)}
    if not output["writable"]:
        errors.append("E06 输出目录不可写")

    level = "red" if errors else ("yellow" if warnings else "green")
    result = {
        "level": level,
        "input": input_info,
        "structure": structure,
        "tools": tools,
        "fonts": fonts,
        "template": tmpl,
        "output": output,
        "warnings": warnings,
        "errors": errors,
        "notes": notes,
    }

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("预检结果：%s" % level)
        print("输入：%s（%s，%s 字节）" % (src.name, ext or "无扩展名", input_info["size"]))
        if structure:
            print("结构：段落 %s，表格 %s，图片 %s，批注 %s，修订 %s" % (
                structure["paragraphs"], structure["tables"], structure["images"],
                structure["comments"], structure["revisions"]))
            c = structure["complex"]
            print("复杂结构：文本框 %s，图形 %s，图片对象 %s，嵌入对象 %s，合并单元格 %s" % (
                c["text_boxes"], c["drawings"], c["pictures"], c["objects"], c["merged_cells"]))
        exact = [name for name, info in fonts["targets"].items() if info["status"] == "exact"]
        parts = []
        if exact:
            parts.append("精确 " + "、".join(exact))
        if fonts["fallbacks"]:
            parts.append("回退 " + "、".join(fonts["fallbacks"]))
        if fonts["missing"]:
            parts.append("缺失 " + "、".join(fonts["missing"]))
        print("字体：%s" % ("；".join(parts) if parts else "未检测"))
        print("模板：%s" % ("可用" if tmpl["exists"] and not tmpl.get("styles_missing") else "不可用或样式不全"))
        for w in warnings:
            print("提示：" + w)
        for e in errors:
            print("错误：" + e)
        for n in notes:
            print("说明：" + n)


if __name__ == "__main__":
    main()
