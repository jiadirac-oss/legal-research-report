============================================================
 法律 Word 文档生成 Skill —— 初始化包
============================================================

本包包含生成两类法律 Word 文档所需的全部资产与初始化脚本，
可整体拷贝到新电脑后一键部署运行。

【包含内容】
  SKILL.md                    # opencode skill 定义（全局安装入口）
  install.ps1                 # 一键初始化脚本
  assets/
    readme法律研究报告.txt     # 法律研究报告完整工作指引（格式参数速查、案例检索流程）
    readme类案检索报告.txt     # 类案检索报告完整工作指引（五C格式参数速查、迭代规范）
    骨架库/
      法律研究参考骨架/        # 法律研究报告 docx 骨架（settings去rsid、document.xml空body保留命名空间头）
      类案检索报告参考骨架/    # 类案检索报告 docx 骨架（含 header/footer rId9~rId17 关系）
    历史版本/                 # 历史 readme 备份（回退用）

【在新电脑上的安装步骤】
  1. 安装 opencode（若未安装）。macOS 还需安装 PowerShell（pwsh）。
  2. 将本目录整体拷贝到新电脑任意位置。
  3. 运行安装脚本（选其一）：
     Windows:
       powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
     macOS:
       pwsh -NoProfile -File install.ps1
     如需自定义部署目录，加参数：
       install.ps1 -WorkDir "D:\我的类案目录"   （Windows）
       install.ps1 -WorkDir "/Users/me/我的类案目录" （macOS）
  4. 脚本自动完成：
      ① 安装全局 skill 到 ~/.config/opencode/skills/law-doc-generation/
      ② 部署骨架库 + readme 到工作目录（默认 桌面\实习VIBE\类案检索报告 或 ~/Desktop/实习VIBE/类案检索报告）
      ③ 生成/合并北大法宝 MCP 配置到 ~/.config/opencode/opencode.jsonc
      ④ 验证 Word 可用（Windows 用 COM；macOS 检查 /Applications/Microsoft Word.app）
  5. 退出并重启 opencode，使 skill 与 MCP 生效。

【使用】
  重启 opencode 后，新会话中说出以下任意关键词即可自动加载本 skill：
    - “生成法律研究报告”
    - “生成类案检索报告”
    - “给类案检索报告加案例 / 重排案例 / 重编号 / 更新页码”
  skill 加载后会先读取对应 readme，再按 readme 规范生成/迭代 Word 文档，
  全程不依赖任何模板 docx。

【注意事项】
  - 本包中的 MCP 配置含北大法宝 API key，请勿公开分享本包。
  - 若已存在 ~/.config/opencode/opencode.jsonc 且含 mcp 配置，脚本不会覆盖，
    需手动合并北大法宝的 mcp 段。
  - macOS 上没有 Word COM：生成后的页码/版面验证需在 Word 中打开人工确认，
    或使用 AppleScript（osascript）读取；脚本第 4 步仅检查 Word 是否已安装。

【与本项目现有文件的关系】
  本包 = 项目目录下 readme + 骨架库 + 历史版本的打包快照。
  日常使用直接以项目目录中的文件为准；本包用于迁移/初始化新电脑。
