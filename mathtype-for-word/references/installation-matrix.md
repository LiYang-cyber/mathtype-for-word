# Installation matrix

## Prerequisites shared by all agents

- Windows 10 or 11.
- Microsoft Word and PowerPoint desktop with COM automation.
- Desktop **MathType for Windows** with its Word template and PowerPoint add-in enabled. This repo is verified against **MathType-win-zh-7.11.1.462** (`ProductVersion 7.11.1.462`) and MathType 7.8.2.441. Download it from <https://mathtype.tw/download/>.
- Do not substitute **MathType Add-In for Microsoft 365**. The task-pane add-in lacks the desktop OLE/COM integration required here.
- Python 3 available as `python.exe` or `python` on `PATH`.
- PowerShell 7 or later, available as Windows `pwsh.exe`. Windows PowerShell 5.1 is not supported because MathType COM macro calls can block under that host.
- LibreOffice (optional) for `preview` PDF/PNG rendering.

## Terminal selection

Detect the active terminal and `pwsh.exe` before changing agent configuration:

| Environment | Action |
|---|---|
| PowerShell 7+ | Run the bridge through the active `pwsh.exe`. |
| Bash on Windows, including Git Bash | Invoke Windows `pwsh.exe` with forward-slash paths. |
| CMD | Invoke `pwsh.exe` with the same arguments. |
| Windows PowerShell 5.1 | Leave 5.1, open Git Bash or CMD, and invoke Windows `pwsh.exe`. |
| WSL Bash | Invoke Windows `pwsh.exe`; Linux `pwsh` cannot drive Windows Office COM. |
| No supported terminal or no `pwsh.exe` | Stop and prompt the user to install PowerShell 7 using <https://learn.microsoft.com/zh-tw/powershell/scripting/install/microsoft-update-faq?view=powershell-7.6>. |

CMD and Git Bash are terminal hosts, not replacements for the PowerShell 7 runtime. A PowerShell 5.1 machine still requires `pwsh.exe` before this bridge can run.

## Repository layout

Set `<REPO_ROOT>` to the absolute path of the checked-out repository before using the commands below.

```text
<REPO_ROOT>\
  README.md
  LICENSE
  scripts\install-skill.ps1
  mathtype-for-word\              the skill payload
    SKILL.md
    LICENSE.txt
    agents\openai.yaml
    config\defaults.json
    examples\example-manifest.json
    references\*.md
    scripts\mathtype-word.ps1            the bridge
    scripts\mathtype-word-editing.ps1    inventory, editing, preview, discovery
    scripts\invoke-bridge.ps1            wrapper that resolves the bridge next to itself
    scripts\mcp_server.py                stdio MCP server
    scripts\run-mcp.ps1
```

Run the probe before installation claims are made:

```console
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\mathtype-for-word\scripts\invoke-bridge.ps1" -Action probe
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\mathtype-for-word\scripts\invoke-bridge.ps1" -Action probe-pptx
```

## Standalone skill directory (Codex, Claude Code, DSH, agents)

Codex discovers skills in `%USERPROFILE%\.codex\skills\<name>\`, Claude Code in `%USERPROFILE%\.claude\skills\<name>\`, and the agent harnesses that read the shared agents directory in `%USERPROFILE%\.agents\skills\<name>\`. All of them expect `SKILL.md` at the root of the skill directory.

The skill payload here is already self-contained: the bridge, its editing module, the MCP server, the default profile, and the references all live inside `mathtype-for-word\`. Copy it, or let the installer do it:

```powershell
# The conventional user roots for Codex, shared agents, and DSH
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\install-skill.ps1" -DefaultRoots

# Or an explicit directory that holds skills (Claude Code, a project, anywhere)
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\scripts\install-skill.ps1" -SkillRoot "$env:USERPROFILE\.claude\skills"
```

Installed layout:

```text
<skill-root>\mathtype-for-word\
  SKILL.md
  LICENSE.txt
  agents\openai.yaml
  config\defaults.json
  examples\example-manifest.json
  references\*.md
  scripts\mathtype-word.ps1
  scripts\mathtype-word-editing.ps1
  scripts\invoke-bridge.ps1
  scripts\mcp_server.py
  scripts\run-mcp.ps1
  scripts\normalize-word-inline-math-layout.py
```

Run the bridge from an installed skill with the wrapper:

```powershell
pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File <skill-root>\mathtype-for-word\scripts\invoke-bridge.ps1 -Action probe
```

MathType Commands does **not** have to be installed into Word. `probe_mathtype_word` discovers the template inside the MathType installation and the bridge loads it as a Word add-in for the duration of each run. `install-template` (or `configure-defaults -InstallTemplate`) copies it into `%APPDATA%\Microsoft\Word\STARTUP` when a permanent Word-side install is wanted.

## Codex

Register and inspect the local MCP server:

```console
codex mcp add mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\mathtype-for-word\scripts\run-mcp.ps1"
codex mcp get mathtype-for-word
```

Restart Codex after changing installed skills.

## Claude Code

Register and inspect the local MCP server:

```console
claude mcp add --scope user mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\mathtype-for-word\scripts\run-mcp.ps1"
claude mcp get mathtype-for-word
```

Restart Claude Code after installation.

## Claude Desktop

Claude Desktop needs both layers:

1. In **Customize > Skills**, upload a ZIP of the `mathtype-for-word` folder so Claude can follow the workflow. See [Use Skills in Claude](https://support.claude.com/en/articles/12512180-use-skills-in-claude).
2. Merge the following entry into `%APPDATA%\Claude\claude_desktop_config.json`. Do not overwrite unrelated existing servers.

Equivalent server entry:

```json
{
  "mcpServers": {
    "mathtype-for-word": {
      "command": "pwsh.exe",
      "args": [
        "-NoLogo",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        "<REPO_ROOT>\\mathtype-for-word\\scripts\\run-mcp.ps1"
      ]
    }
  }
}
```

Restart Claude Desktop, call the probe tool, and then run a small test document before using a valuable one.

## ChatGPT Desktop

ChatGPT Desktop does not automatically discover a local repository or Codex skill directory. It requires a remote MCP endpoint or OpenAI's Secure MCP Tunnel, routed back to the signed-in interactive Windows desktop where Office COM can run. This repository ships the local stdio server only.

## Verification after installation

1. List the skill and MCP tool set in the host.
2. Call `probe_mathtype_word` and `probe_mathtype_powerpoint`.
3. Render and validate one small DOCX and one small PPTX of your own, then inspect the rendered pages.
4. Require `ok: true` from every `validate` call.
5. Open the DOCX in Word and confirm that double-clicking the equation opens MathType and the reference navigates to the number.
6. Open the PPTX in PowerPoint and confirm that editing the floating equation opens desktop MathType.
