# 法律 Word 文档生成 Skill
来自于某红圈所实习生
一个面向法律实务的 [opencode](https://opencode.ai) skill 包，用于自动生成与迭代两类法律 Word 文档：

- **法律研究报告** —— 初步结论 + 法规梳理 + 理论观点 + 类案支撑
- **类案检索报告** —— 封面 + 检索说明 + 检索结论 + 案例摘要比对表 + 案例原文附件

所有格式规则、精确 OOXML 参数、生成步骤均固化在资产文件中，生成全程**不依赖任何模板 docx**，直接在 docx 骨架的 `document.xml` 中构造正文后重新打包。

## 目录结构

```
law-doc-generation/
├── SKILL.md                  # opencode skill 定义（全局安装入口）
├── README.txt                # 初始化包说明与安装步骤
├── install.ps1               # 一键初始化脚本（Windows PowerShell / macOS pwsh）
└── assets/
    ├── readme法律研究报告.txt # 法律研究报告完整工作指引（格式参数速查、案例检索流程）
    ├── readme类案检索报告.txt # 类案检索报告完整工作指引（五C格式参数速查、迭代操作规范）
    ├── 骨架库/
    │   ├── 法律研究参考骨架/   # 法律研究报告 docx 骨架
    │   └── 类案检索报告参考骨架/ # 类案检索报告 docx 骨架（含 header/footer 关系）
    └── 历史版本/              # 历史 readme 备份（回退用，可忽略）
```

## 安装

1. 安装 [opencode](https://opencode.ai)；macOS 另需安装 PowerShell（`pwsh`）。
2. 将本仓库整体克隆/拷贝到目标电脑。
3. 运行初始化脚本：

   **Windows**
   ```powershell
   powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
   ```

   **macOS**
   ```bash
   pwsh -NoProfile -File install.ps1
   ```

   自定义部署目录：
   ```powershell
   install.ps1 -WorkDir "D:\我的类案目录"                    # Windows
   install.ps1 -WorkDir "/Users/me/我的类案目录"              # macOS
   ```

4. 脚本自动完成：
   1. 安装全局 skill 到 `~/.config/opencode/skills/law-doc-generation/`
   2. 部署骨架库 + readme 到工作目录
   3. 生成/合并北大法宝 MCP 配置到 `~/.config/opencode/opencode.jsonc`
   4. 验证 Word 可用（Windows 用 COM；macOS 检查 Word 是否安装）

5. 重启 opencode，使 skill 与 MCP 生效。

## 使用

重启 opencode 后，在新会话中说出以下任意关键词即可自动加载本 skill：

- 「生成法律研究报告」
- 「生成类案检索报告」
- 「给类案检索报告加案例 / 重排案例 / 重编号 / 更新页码」

skill 加载后会先读取对应 readme 工作指引，再按其中的强制规则生成/迭代 Word 文档。

## 注意事项

- **API key 已脱敏**：`install.ps1` 中的北大法宝 MCP 配置密钥已替换为 `<YOUR_PKULAW_API_KEY>` 占位符，安装前请自行填入真实密钥。
- 若本机已存在 `~/.config/opencode/opencode.jsonc` 且含 mcp 配置，脚本不会覆盖，需手动合并北大法宝的 mcp 段。
- macOS 没有 Word COM：生成后的页码/版面验证需在 Word 中人工确认，或使用 AppleScript（`osascript`）读取。
- 类案检索报告的案例排序规则：指导性/入库/参考案例置前，其后按最高人民法院 → 高级人民法院 → 中级人民法院 → 基层人民法院，同级保持相对顺序。

## 平台适配

| 环节 | Windows | macOS |
| --- | --- | --- |
| 初始化脚本 | `powershell` | `pwsh` |
| Word 验证 | Word COM | AppleScript / 人工确认 |
| 文档生成 | 相同（docx 骨架 + document.xml 构造） | 相同 |

## 许可

[MIT](LICENSE)
