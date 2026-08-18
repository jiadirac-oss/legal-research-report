---
name: law-doc-generation
description: 生成与迭代两类法律 Word 文档——"法律研究报告"和"类案检索报告"。当用户要求生成法律研究报告、类案检索报告、或对既有报告增删/重排案例、重编号、更新页码时使用。触发关键词：法律研究报告、类案检索报告、初步结论、检索结论、案例摘要比对表、附件案例、封面、检索情况的说明、类案。
---

# 法律 Word 文档生成（法律研究报告 + 类案检索报告）

本 skill 指导如何生成两类法律 Word 文档：**法律研究报告**（初步结论＋法规＋理论观点＋案例）与**类案检索报告**（封面＋检索说明＋检索结论＋案例摘要比对表＋案例原文附件）。所有格式规则、精确 OOXML 值、生成步骤均已固化在资产文件的两个 readme 中，生成时**直接依据 readme 执行，无需再打开任何模板 docx**。

## 资源文件（本 skill 目录内）

```
SKILL.md                            # 本文件
assets/
  readme法律研究报告.txt            # 法律研究报告的完整工作指引（格式参数速查、案例检索流程）
  readme类案检索报告.txt            # 类案检索报告的完整工作指引（五C格式参数速查、迭代操作规范）
  骨架库/
    法律研究参考骨架/                # 法律研究报告 docx 骨架（settings去rsid、document.xml空body保留命名空间头）
    类案检索报告参考骨架/            # 类案检索报告 docx 骨架（同上，含 header/footer rId9~rId17 关系）
  历史版本/                         # 历史 readme 版本（回退用，可忽略）
```

## 使用流程

1. **先读 readme**：开始工作前，先读取对应的 readme 全文：
   - 生成/迭代"法律研究报告" → 读 `assets/readme法律研究报告.txt`
   - 生成/迭代"类案检索报告" → 读 `assets/readme类案检索报告.txt`
2. **遵循 readme 内置规则**（均为强制要求）：
   - 【双向钢人论证规则】【执行前提问规则】【历史版本规则】（更新 readme 前必须备份历史版本）。
   - 法律研究报告：按【板块一~五】【案例检索完整流程】【格式参数速查】执行。
   - 类案检索报告：按【二~五C 结构与格式】【六 案例检索要点】【七 迭代操作规范】执行；案例排序规则为"指导/入库/参考案例置前 → 最高院→高院→中院→基层，同级保持相对顺序"。
3. **生成 Word（不依赖模板 docx）**：
   - 复制对应骨架库目录 → 替换其中 `word/document.xml` 的 `<w:body>...</w:body>`（保留 `<w:body>` 前的全部命名空间声明头与 `</w:body>` 后的收尾）→ 重新压缩为 docx。
   - 法律研究骨架 body 末尾需自加第 1 节 sectPr（footerReference rId14，pgMar top/bottom=1440 left/right=1800，pgNumType start=1，docGrid linePitch=312）。
   - 类案检索骨架 body 需整体构造：封面24段（含第1节sectPr rId9/rId10）→ 检索说明/结论/表格一 → 结尾（第2节sectPr rId11~rId16）→ 附件案例块 → 第3节sectPr（rId17）。3 个 sectPr 精确值见 readme【五C 六】。
   - 编码：document.xml 以 UTF-8（无BOM）写入；脚本须 UTF-8 带 BOM 保存，PowerShell 7.7 实战教训见 readme 类案检索报告【7.7】。
4. **验证**：生成后用 Word 打开确认段落数/表格数/节数/页数符合预期：
   - Windows：可用 Word COM（`New-Object -ComObject Word.Application`）读取；类案检索报告按 readme【输出要求】用 Word COM 精确读取案例起始页计算表格一"页码"列。
   - macOS：无 Word COM，改用 AppleScript（`osascript -e 'tell application "Microsoft Word"...'`）或直接打开 docx 在 Word 中人工确认段落/表格/节/页数；页码列需在 Word 中查看分页后人工回填或核对。
5. **迭代修改**（类案检索报告）：不替换 body，只增删改 body 内段落/表格；3 个 sectPr 与页眉页脚关系保持不变；重排后必须重编号案例标题与表格一序号、重算页码。

## 平台适配说明（Windows / macOS）

- **工作目录**：readme 中出现的绝对路径（如 `C:\Users\jiajinhua\Desktop\实习VIBE\类案检索报告\`）为 Windows 开发机示例。实际执行时以"当前会话的工作目录"为准——install.ps1 会把骨架库部署到该目录（macOS 默认 `~/Desktop/实习VIBE/类案检索报告`），生成时从该目录下的 `骨架库\` 取骨架。
- **Word 验证**：Windows 用 Word COM；macOS 无 COM，用 AppleScript（osascript）或手动在 Word 中打开确认。页码计算需在 Word 中读取真实分页。
- **脚本**：install.ps1 兼容 Windows（`powershell`）与 macOS（`pwsh`）。PowerShell 中文/编码规则对两者同样适用。

## 部署与初始化

- 本 skill 已安装至全局 `~/.config/opencode/skills/law-doc-generation/`，任何目录的新会话均可用。
- 若需在新电脑整体部署，运行 `install.ps1`（位于 skill 包根目录或 skill 目录内）：它会 ①安装/更新全局 skill ②把骨架库+readme 部署到工作目录 ③生成北大法宝 MCP 配置 ④验证 Word 可用（Windows COM / macOS 检查 Word 安装）。
