# ============================================================
#  法律 Word 文档生成 skill —— 初始化安装脚本
#  功能：
#    1. 安装/更新全局 skill 到 ~/.config/opencode/skills/law-doc-generation/
#    2. 把骨架库 + readme 部署到工作目录（默认桌面 实习VIBE\类案检索报告）
#    3. 生成/合并北大法宝 MCP 配置到 ~/.config/opencode/opencode.jsonc
#    4. 验证 Word 可用（Windows 用 COM 探测；macOS 检查 /Applications/Microsoft Word.app）
#  用法：
#    Windows: powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
#    macOS:   pwsh -NoProfile -File install.ps1
#    可选参数：-WorkDir "D:\我的类案目录" 或 -WorkDir "/Users/me/我的类案目录"
# ============================================================
[CmdletBinding()]
param(
    [string]$WorkDir = $(if ($IsMacOS) { "$HOME/Desktop/实习VIBE/类案检索报告" } else { "$HOME\Desktop\实习VIBE\类案检索报告" })
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$u8 = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName System.IO.Compression.FileSystem

$ScriptRoot = $PSScriptRoot
$SkillName = "law-doc-generation"
$HomeDir = if ($IsMacOS) { $HOME } else { $env:USERPROFILE }
$TempDir = [System.IO.Path]::GetTempPath()
$SkillDir = Join-Path $HomeDir ".config/opencode/skills/$SkillName"
$GlobalSkillDir = Join-Path $HomeDir ".config/opencode/skills"

function Step($msg) { Write-Host "`n== $msg ==" -ForegroundColor Cyan }
function OK($msg)   { Write-Host "  [OK] $msg" -ForegroundColor Green }
function Warn($msg){ Write-Host "  [!] $msg" -ForegroundColor Yellow }
function Fail($msg){ Write-Host "  [X] $msg" -ForegroundColor Red }

# ---------- locate assets ----------
$Assets = if (Test-Path (Join-Path $ScriptRoot "assets")) { Join-Path $ScriptRoot "assets" }
          elseif (Test-Path (Join-Path $ScriptRoot "..\assets")) { Join-Path $ScriptRoot "..\assets" }
          else { $null }
if (-not $Assets) { Fail "未找到 assets 目录，请确保 install.ps1 与 assets/ 位于同一目录（或上一级）"; exit 1 }

$LrSkeleton = Join-Path $Assets "骨架库/法律研究参考骨架"
$CaSkeleton = Join-Path $Assets "骨架库/类案检索报告参考骨架"
$LrReadme   = Join-Path $Assets "readme法律研究报告.txt"
$CaReadme   = Join-Path $Assets "readme类案检索报告.txt"

# ============================================================
# 1. install / update global skill
# ============================================================
Step "1/4 安装全局 skill"
New-Item -ItemType Directory -Force -Path $GlobalSkillDir | Out-Null
New-Item -ItemType Directory -Force -Path $SkillDir | Out-Null
if (Test-Path $SkillDir) { Remove-Item $SkillDir -Recurse -Force; New-Item -ItemType Directory -Force -Path $SkillDir | Out-Null }
Copy-Item (Join-Path $ScriptRoot "SKILL.md") (Join-Path $SkillDir "SKILL.md")
New-Item -ItemType Directory -Force -Path (Join-Path $SkillDir "assets/骨架库") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $SkillDir "assets/历史版本") | Out-Null
Copy-Item $LrReadme                 (Join-Path $SkillDir "assets/")
Copy-Item $CaReadme                 (Join-Path $SkillDir "assets/")
Copy-Item $LrSkeleton               (Join-Path $SkillDir "assets/骨架库/") -Recurse
Copy-Item $CaSkeleton               (Join-Path $SkillDir "assets/骨架库/") -Recurse
if (Test-Path (Join-Path $Assets "历史版本")) {
    Copy-Item (Join-Path $Assets "历史版本/*") (Join-Path $SkillDir "assets/历史版本/") -Recurse
}
OK "skill 已安装到 $SkillDir"

# ============================================================
# 2. deploy skeleton + readme to work dir
# ============================================================
Step "2/4 部署骨架库与 readme 到工作目录"
New-Item -ItemType Directory -Force -Path (Join-Path $WorkDir "骨架库") | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $WorkDir "历史版本") | Out-Null
Copy-Item $LrReadme $WorkDir -Force
Copy-Item $CaReadme $WorkDir -Force
Copy-Item $LrSkeleton (Join-Path $WorkDir "骨架库/") -Recurse -Force
Copy-Item $CaSkeleton (Join-Path $WorkDir "骨架库/") -Recurse -Force
OK "骨架库与 readme 已部署到 $WorkDir"

# ============================================================
# 3. generate / merge 北大法宝 MCP config
# ============================================================
Step "3/4 配置北大法宝 MCP"
$OcDir = Join-Path $HomeDir ".config/opencode"
$OcFile = Join-Path $OcDir "opencode.jsonc"
New-Item -ItemType Directory -Force -Path $OcDir | Out-Null

$mcpBlock = @"
    "mcp": {
        "law_recognition":        { "type": "remote", "url": "https://apim-gateway.pkulaw.com/law_recognition",        "headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "mcp-fatiao":             { "type": "remote", "url": "https://apim-gateway.pkulaw.com/mcp-fatiao",             "headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "mcp-law":                { "type": "remote", "url": "https://apim-gateway.pkulaw.com/mcp-law",                "headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "mcp-law-search-service": { "type": "remote", "url": "https://apim-gateway.pkulaw.com/mcp-law-search-service", "headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "add-doc-link":           { "type": "remote", "url": "https://apim-gateway.pkulaw.com/add-doc-link",           "headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "pku_citation_validator": { "type": "remote", "url": "https://apim-gateway.pkulaw.com/pku_citation_validator", "headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "case_number_recognition":{ "type": "remote", "url": "https://apim-gateway.pkulaw.com/case_number_recognition","headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "mcp-case":               { "type": "remote", "url": "https://apim-gateway.pkulaw.com/mcp-case",               "headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "mcp-law-agg":            { "type": "remote", "url": "https://apim-gateway.pkulaw.com/mcp-law-agg",            "headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } },
        "mcp-case-search-service":{ "type": "remote", "url": "https://apim-gateway.pkulaw.com/mcp-case-search-service","headers": { "Authorization": "Bearer <YOUR_PKULAW_API_KEY>" } }
    }
"@

$fullConfig = @"
{
  "`$schema": "https://opencode.ai/config.json",
  "tool_output": {
    "max_lines": 5000,
    "max_bytes": 300000
  },
  "experimental": {
    "mcp_timeout": 60000
  },
$mcpBlock
}
"@

if (Test-Path $OcFile) {
    $old = [System.IO.File]::ReadAllText($OcFile, $u8)
    if ($old -match '"mcp"\s*:\s*\{') {
        Warn "opencode.jsonc 已包含 mcp 配置，跳过覆盖（如需更新请手动编辑）。位置: $OcFile"
    } else {
        $bak = "$OcFile.bak_" + (Get-Date -Format "yyyyMMdd_HHmmss")
        Copy-Item $OcFile $bak -Force
        [System.IO.File]::WriteAllText($OcFile, $fullConfig, $u8)
        OK "已备份原配置至 $bak 并写入北大法宝 MCP 配置"
    }
} else {
    [System.IO.File]::WriteAllText($OcFile, $fullConfig, $u8)
    OK "已创建 $OcFile 并写入北大法宝 MCP 配置"
}

# ============================================================
# 4. verify Word availability
# ============================================================
if ($IsMacOS) {
    Step "4/4 检查 Microsoft Word"
    $wordApp = "/Applications/Microsoft Word.app"
    if (Test-Path $wordApp) { OK "已检测到 Microsoft Word（$wordApp）。macOS 上无 Word COM，生成后的验证改用 AppleScript 或手动打开确认。" }
    else { Warn "未在 /Applications 找到 Microsoft Word.app。macOS 无 Word COM，生成后的页码/版面验证需在 Word 中手动打开确认。" }
} else {
    Step "4/4 验证 Word COM"
    # 在独立子进程中做 COM 探测，带硬超时，防止 Word 弹窗导致安装脚本挂起
    $probe = Join-Path $TempDir ("skill_probe_" + (Get-Random) + ".ps1")
    $probeSrc = @'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Get-Process -Name WINWORD,WPS,wpsoffice -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
$testDocx = Join-Path $env:TEMP ("skill_verify_" + (Get-Random) + ".docx")
try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0
    $doc = $w.Documents.Add()
    $doc.Content.Text = "skill verify ok"
    $doc.SaveAs2($testDocx, 16)
    $doc.Close(0)
    $w.Quit()
    Write-Output ("WORD-OK " + (Test-Path $testDocx))
} catch {
    Write-Output ("WORD-FAIL " + $_.Exception.Message)
}
'@
    [System.IO.File]::WriteAllText($probe, $probeSrc, (New-Object System.Text.UTF8Encoding($true)))
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$probe`"" -PassThru -WindowStyle Hidden -Wait:$false
    if (-not $proc.WaitForExit(60000)) {
        $proc.Kill()
        Warn "Word COM 验证超时（60 秒），已跳过。请确认已安装 Microsoft Word 或 WPS。"
    } else {
        $out = $proc.StandardOutput.ReadToEnd()
        if ($out -match "WORD-OK") { OK "Word COM 可用（$($out.Trim())）" }
        else { Warn "Word COM 验证未通过：$($out.Trim())" }
    }
    Remove-Item $probe -Force -ErrorAction SilentlyContinue
}

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "  安装完成。请退出并重启 opencode 使 skill 与 MCP 生效。" -ForegroundColor Yellow
Write-Host "  新会话中说出“生成法律研究报告”或“生成类案检索报告”即可触发本 skill。" -ForegroundColor Yellow
Write-Host "========================================================" -ForegroundColor Cyan