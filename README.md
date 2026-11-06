# MathType for Word（中文整合版）

把 TeX 公式写成真正的 **MathType 7 Equation.DSMT4** 对象，在 Word DOCX / PowerPoint PPTX 里排版、编号、交叉引用，并就地编辑文档里已有的公式。全程走桌面版 Word/PowerPoint 的 COM 自动化，不生成 Word 自带 OMath、公式图片，也不用题注文字冒充公式。

本仓库是 [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word) 的分支，保留上游 MIT 许可与署名。上游面向 Codex / Claude 插件市场，内置批量渲染、纯文本公式扫描打标、PowerShell 7 自动安装和测试套件；本分支面向「单技能目录 + 中文文档 + 就地编辑已有公式」的场景。

## 与上游的关系

- 基线：上游 `441daf4`（2026-07-22，1.3.x 时代代码）。
- 上游在 2026-10-13 发布的 1.4.0 重写了桥脚本（55 KB → 105 KB），新增 `scan` / `mark`、批量渲染与断点续跑、桥锁、孤儿 Word 清理、PowerShell 7 自动安装，以及插件市场打包、`tests/`、`evals/`、CI。
- 本分支没有合入上述能力，改为在旧桥脚本上增加就地编辑与中文学术排版支持。

## 本分支新增

| 能力 | 说明 |
|---|---|
| 就地编辑已有公式 | `inventory` 列出每行公式的序数、布局与缓存 TeX；`edit` 按 `old_tex` 校验后原地重写，复用原定界符，支持删除与正文精确替换，校验失败不发布 |
| 公式预览 | `preview` 用 LibreOffice 把 DOCX 转成 PDF 与逐页 PNG，供目检排版 |
| 中文学术排版 | `其中，` 段零首行缩进、内联公式 `w:position=0` 与 `w:textAlignment=center`、禁止固定行距溢出；`normalize-word-inline-math-layout.py --check` 自动检查 |
| Texvc 实测表 | MathType 7.8 上哪些 TeX 命令变红字、黑菱形或静默丢符号，附等价替换写法 |
| 环境探测增强 | MathType 从 PATH、注册表、安装根目录发现；Word 模板按需加载；用户目录可写性与 Vector 样式检查 |
| 单技能安装 | 桥脚本、编辑模块、MCP server 都在 `mathtype-for-word\` 内，复制即可用，无需插件根目录或环境变量 |

## 本分支未包含（相对上游）

- `scan` / `mark`：从纯文本公式扫描候选并打标；
- 批量渲染、断点续跑、后台任务状态查询；
- 桥锁、孤儿 Word 清理、MathType 偏好 journal、行距裁剪修复；
- PowerShell 7 自动安装（要求本机已有 `pwsh.exe`）；
- 插件市场打包、`tests/`、`evals/`、CI。

超过 40 个公式的长文档建议使用上游版本；本分支走单 Word 会话渲染。

## 目录

| 路径 | 作用 |
|---|---|
| `mathtype-for-word/SKILL.md` | 技能契约：工作流、校验门、编号与引用策略、失败处理 |
| `mathtype-for-word/scripts/mathtype-word.ps1` | Word / PowerPoint 的 COM 桥，动作见下 |
| `mathtype-for-word/scripts/mathtype-word-editing.ps1` | 就地编辑模块，被桥脚本 dot-source 加载 |
| `mathtype-for-word/scripts/mcp_server.py`、`run-mcp.ps1` | 本机 stdio MCP server |
| `mathtype-for-word/scripts/normalize-word-inline-math-layout.py` | 内联公式排版检查与原地修正 |
| `mathtype-for-word/references/` | 学术公式排版、编号与引用、PPT 工作流、安装矩阵、排错、manifest 结构 |
| `scripts/install-skill.ps1` | 把技能安装到 Codex / 共享 agents / DSH 的技能根目录 |

桥脚本的动作：`probe`、`probe-pptx`、`configure-defaults`、`install-template`、`inventory`、`render`、`validate`、`edit`、`update`、`preview`、`render-pptx`、`validate-pptx`、`cleanup`。

## 前置条件

- Windows 10 或 11，Microsoft **Word / PowerPoint 桌面版**。Microsoft 365 的网页式 MathType Add-in 不行，本工具依赖桌面 OLE 与 Word 模板。
- **MathType for Windows**（上游以 7.11.1.462 开发验证，本分支另在 7.8.2.441 上实测）。只装 Add-in 不够。
- PowerShell 7（`pwsh`）；MCP server 另需 PATH 上有 Python 3。
- 可选：LibreOffice，用于 `preview` 生成 PDF/PNG。

动手前先跑 `probe`：只读，报告 MathType、Word 模板、用户目录可写性、Vector 样式等检查项；失败就停，不要硬跑。

## 安装

1. **作为技能安装（推荐）**

   ```powershell
   pwsh -File scripts\install-skill.ps1 -DefaultRoots
   ```

   默认写入 `%USERPROFILE%\.codex\skills\`、`%USERPROFILE%\.agents\skills\`、`%USERPROFILE%\.dsh\skills\`；也可用 `-SkillRoot` 指定任意技能根目录，用 `-Force` 覆盖已有安装。目标目录里的技能名固定为 `mathtype-for-word`。

   手工安装等价于把 `mathtype-for-word\` 整个目录复制到技能根目录，文件夹名保持不变。

2. **注册 MCP server**

   ```powershell
   codex mcp add mathtype-for-word -- pwsh.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "<REPO_ROOT>\mathtype-for-word\scripts\run-mcp.ps1"
   ```

   完整矩阵见 [`mathtype-for-word/references/installation-matrix.md`](mathtype-for-word/references/installation-matrix.md)。

3. **直接跑桥脚本**

   ```powershell
   pwsh -File mathtype-for-word\scripts\mathtype-word.ps1 -Action probe
   ```

   `render` / `validate` / `edit` 需要 `-InputPath`、`-OutputPath`、`-ManifestPath` 等参数，完整契约见 `SKILL.md` 与 `references/workflow-and-schema.md`。

## 边界与常见坑

- 原始输入只读，输出另给路径；原地编辑需要显式授权 `-Overwrite`。
- 一律生成可编辑 MathType 对象，不用 OMath、公式图片或普通文字替代。
- 受限沙箱里 Word COM 会报 `Word could not fire event`（`0x800A1772`），原因是写不了 `%APPDATA%\Microsoft\Word`；需要在全盘权限下运行。
- 复杂分式、矩阵、对齐组、重音、嵌套上下标必须渲染成 PDF/PNG 逐页看，结构校验证明不了排版正确。
- 编号与交叉引用由 Word 原生域承担，移动或增删公式后要跑 `update` 更新域。
- 插入右侧编号时，桥脚本用临时书签区分 MathType 新增的空段落与用户原有空行，只清理前者。
- 右编号公式后的段落以「其中，」开头时，自动取消首行缩进，并让前导正文、公式和说明同页衔接。

## 版本与许可

- 本分支版本 **2.0.0**（2026-11-06），基于上游 `441daf4`，MIT 许可。
- 上游项目：<https://github.com/felimet/mathtype-for-word>
- 版权：上游 `LICENSE` 保留 Jia-Ming Zhou（Felimet）署名；本分支的修改同样以 MIT 发布。
