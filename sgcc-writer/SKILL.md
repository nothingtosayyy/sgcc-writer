---
name: sgcc-writer
name_en: SGCC Writer
name_zh: 国网文档助手
description: Write and revise State Grid (国网) Word materials — plans, reports, summaries, standards, research — and apply SGCC house formatting to any Word document on request (排版、统一格式、套模板、格式规范化). Auto-use for 国网/国家电网/省公司 content or explicit formatting; revision uses targeted edits with a change log; formatting does not change content. Word only; PPT and speech drafts are out of scope.
description_en: Write and revise State Grid (国网) Word materials — plans, reports, summaries, standards, research — and apply SGCC house formatting to any Word document on request (排版、统一格式、套模板、格式规范化). Auto-use for 国网/国家电网/省公司 content or explicit formatting; revision uses targeted edits with a change log; formatting does not change content. Word only; PPT and speech drafts are out of scope.
description_zh: 国网 WORD 材料写作、改稿与排版技能。覆盖方案、书面汇报（含总结）、标准规范、调研报告；自动适用国网/国家电网/省公司内容或明确排版请求；改稿定点编辑并输出变更对照，排版不改内容。仅 WORD，PPT、讲话稿不在范围。
when_to_use: 国网 Word 材料写作、改稿、排版、统一格式、套模板、格式规范化，以及润色/顺句/按批注修改/压缩扩写；不用于讲话稿、PPT、纯问答或代码任务。
argument-hint: Describe the Word document to write or format, e.g. "帮我写一个2027年样本运营工作方案" or "把这份汇报按国网格式排一下"
argument-hint-en: Describe the Word document to write or format, e.g. "帮我写一个2027年样本运营工作方案" or "把这份汇报按国网格式排一下"
argument-hint-zh: 描述要写的或要排版的 Word 材料，如"帮我写一个2027年样本运营工作方案"或"把这份汇报按国网格式排一下"
user-invocable: true
---

# 国网文档助手

本技能沉淀国网公文写作惯例（方案、汇报、规范、报告）与国网排版预设，示例均为演示值或虚构值。目标：引导补料 → 按体例生成 → 交付可核验的 docx。

**范围**：WORD 内容写作（方案／汇报含总结／规范／调研）与排版（任意 WORD 文档按国网预设；`.docx` 原生，`.doc` 转换后）。PPT、讲话稿不在范围。

> **免责声明**：本技能中的示例、数字、单位、文号、人名、引文等均为演示值或虚构值，不来自真实文档，不作为事实依据；正式材料中的事实须由用户提供并核实。

## 适用判断

出现以下任一情形即使用本技能：

- 文档内容涉及国网、国家电网、省电力公司、供电公司等单位；
- 文体为国网常见公文（方案、汇报、总结、规范、调研报告），且使用“一、（一）1.（1）”编号链；
- 用户提出润色、顺句、重写章节、按批注修改、压缩扩写、统一口径；
- 用户对任意 WORD 文档明确要求国网排版、统一格式、套模板。

判断存疑时按“疑似国网材料”处理，第一步向用户确认。排版不要求国网特征；无法判断文档类型时按国网预设处理。

## 路径判定

| 场景 | 判定信号 | 走哪条路 |
|---|---|---|
| 从 0 到 1 | 没有现成 WORD 底稿，用户说“写/起草/做一个” | 起草路径（五步工作流） |
| 改这份 | 有现成底稿，用户说“改/润色/顺句/按批注/压缩/扩写/统一口径” | 改稿路径 |
| 参照重写 | 有底稿，但用户要“照这个写一份新的” | 起草路径；原稿只作素材，不输出变更对照 |
| 纯排版 | 有底稿，用户说“排版/统一格式/套模板”，不要求改内容 | 排版路径（reference-format.md） |
| 生成即排版 / 改稿后排版 | 起草完成，或改稿后要求排版 | 按“排版模式”确定依据；改稿后排版默认原稿体例 |

判定只看三件事：**有没有现成底稿**、**用户要改内容还是只排版**、**有没有指定目标格式**。

## 改稿模式

进入条件：有现成 WORD 底稿，用户提出润色、顺句、重写章节、按批注修改、压缩扩写、统一口径。没有点名段落或批注时先问范围，不擅自全文重写。

做法：只改点名/批注位置；沿用原体例、批注锚点和修订痕迹；默认修订模式输出（`w:ins`/`w:del`）；交付附“原→新”变更对照；自检 A+B2+C。

细则、改稿红线与 ins-aware 定位见 [reference-revision.md](reference-revision.md)。

## 排版模式

纯排版只调格式，不改内容。依据三选一：

- **国网预设**：纯排版且未指定目标，一律使用（含非国网文档）；
- **原稿体例**：仅改稿后排版且未指定目标时使用，回执提示可改用国网预设；
- **用户模板**：用户提供模板时使用，缺项用国网预设补齐。

工具按能力类别选择，不绑定名称：本地 Word 文档处理能力（含 Word/WPS 自动化）→ Python + `python-docx` 兜底 → 仅规范/模板。无法自动排版时明确告知，不谎称完成。

开工前按 [reference-runtime.md](reference-runtime.md) 预检；优先套用 `assets/templates/guowang-preset-v1.0.1.docx`，模板不可用再按参数排版。异常按异常卡处理，不覆盖原文件。

交付附排版回执六项：依据、执行方式、已应用、未应用、字体回退、校验结果；另附预检结论、异常/降级。细则、回执模板与 D1–D8 见 [reference-format.md](reference-format.md) 和 [reference-check.md](reference-check.md)。

## 内容与规则分界

- **规则**：体裁骨架、章节功能、编号链、句式骨架、口吻尺度、称谓、命名、自检。
- **内容**：业务专名、指标数值、文号、人名职务、引文原文、具体年度／单位／项目。
- 判据：一条示例换一个主题还能用吗？不能，就是内容。
- `〔 〕`为结构位，成品前必须替换；`【 】`为用户未提供的事实，交付时列清单。
- `examples/语感示例.md` 默认不读；需要校准语感时读，其中内容不得复用。

## 文件版本与归档

- 改前备份到同目录 `历史版本/`；当前目录只留最新稿。
- 起草：`【YYYY.MM.DD】主题+文种_Vn.n`；改稿：原名 + `-YYYYMMDD-vN`；排版：原名 + `-排版-〔依据〕-YYYYMMDD`。
- 改稿默认修订模式；要干净稿时另出接受修订版，不覆盖修订版。
- 用户指定路径、命名或保留多份时，以用户为准。

## 总工作流

```
第一步 判型 → 第二步 引导补料 → 第三步 按模板生成 → 第四步 成稿交付 → 第五步 交稿自检
```

> 改稿走“改稿模式”，纯排版走“排版模式”，均不跑五步工作流。

## 第一步：判型路由

| 用户典型说法 | 文体 | 亚型 | 参考文件 |
|---|---|---|---|
| “写个XX方案”“建设方案” | 工作方案 | 年度／建设专项／落实型／运营／思路举措／流程规范 | [reference-plan.md](reference-plan.md) |
| “写个书面汇报”“给X总写个汇报” | 书面工作汇报 | 标准回顾型／前瞻举措型 | [reference-report.md](reference-report.md) |
| “写个工作总结”“年度总结” | 书面工作汇报 | 总结变体（年度／专项） | [reference-report.md](reference-report.md) §1.1 |
| “起草个规范/标准”“接口规范” | 标准规范 | GB企标体／接口规范体／操作手册体 | [reference-standard.md](reference-standard.md) |
| “调研报告”“研究报告” | 调研报告 | 现状调研／专题研究／举措研究 | [reference-research.md](reference-research.md) |

信号词：出现“文号/主送/印发/遵照执行”属红头公文，不在范围；出现“调研/问题分析/经验总结”走调研报告；出现“全年工作/存在问题/下一年思路”走总结变体；同时要 PPT 时只产出 WORD。

**排版任务不判文体**：用户明确要求排版/统一格式时，直接进入“排版模式”。

## 第二步：引导补料

通用六问（每批≤4 个，未答用【 】）：

1. 主题与事由：材料讲什么工作？
2. 对象与场合：给谁看/听？
3. 时间坐标：什么时间用？
4. 篇幅：给“字数（约几页）”档位——2千字约4页／5千字约9页／1万字约18页。
5. 依据：上级文件文号、领导指示原话、会议要求？
6. 素材：有哪些数据、做法、进展可写入？

页数按 A4／三号仿宋／28 磅、满页约 600 字（含标题按 550 字/页）估算；封面、目录、图表另计。

分文体追问：方案问量化目标、任务维度、分工、节点；汇报问成效数字、问题、下一步；规范问适用范围、应/宜尺度、术语与附录；报告问调研过程、对标对象、建议责任矩阵。

**长文检查点**：预计超过 5000 字（约 9 页）或超过 5 章，先出章节大纲确认。

## 第三步：按模板生成

读对应文体文件 + [reference-style.md](reference-style.md)，严格套用目录骨架、编号体系、句式模板、量化句型和口吻规则。参考文件只写规则与占位骨架；需要语感时再读 examples，且不得复用其中内容。

## 第四步：成稿交付

- **预检**：开工前按 [reference-runtime.md](reference-runtime.md) 检查工具、字体、输入、模板和输出；绿色直接做，黄色先告知限制，红色停止且不写文件。
- **生成**：仅 WORD 类，直接生成 .docx；正文默认按 [reference-format.md](reference-format.md) 的国网预设排版。用户说“先不排版”时只出内容稿。
- **保存**：用户指定路径优先，否则存当前工作目录根目录；中间产物进 `产物/artifacts/`，临时文件进 `产物/tmp/`，脚本进 `产物/scripts/`；收尾登记 `产物/INDEX.md`，只移不删。
- **内容来源清单**：三列 `类别 | 内容 | 来源`，类别覆盖数字、年份、文号、人名职务、引文、专名、提法、单位／场景；来源只填“用户原话／用户文件／【未提供】”；正文每项可溯源，找不到改【 】；不搬用参考文件或 examples 的具体内容。
- **交付回复**：成品绝对路径、字数与页数、未决占位符数量、内容来源清单；涉及排版附回执（依据、执行方式、已应用、未应用、字体回退、校验）。不使用 "outputs" 目录或任何 "present" 操作。

## 第五步：交稿自检

按 [reference-check.md](reference-check.md) 执行：

- **A 共用硬伤**：一票否决；
- **B1 起草体例 / B2 改稿体例**：按路径选一；
- **C 共用交付**；
- **D 排版**：涉及排版时执行。

路径组合：起草 A+B1+C+D；改稿仅内容 A+B2+C；改稿后排版 A+B2+C+D；纯排版 D+C。

交付前检索 `TODO|待补充|待确认`，核对编号跳号、文件名与正文年月。

## 细则索引

- 文风、词汇、句式、称谓、帽段 → [reference-style.md](reference-style.md)
- 方案 → [reference-plan.md](reference-plan.md)；汇报/总结 → [reference-report.md](reference-report.md)
- 规范 → [reference-standard.md](reference-standard.md)；调研 → [reference-research.md](reference-research.md)
- 改稿 → [reference-revision.md](reference-revision.md)；排版 → [reference-format.md](reference-format.md)；自检 → [reference-check.md](reference-check.md)
- 预检、模板、异常 → [reference-runtime.md](reference-runtime.md)；模板文件 → `assets/templates/guowang-preset-v1.0.1.docx`
- 语感样例 → `examples/语感示例.md`（默认不读）
