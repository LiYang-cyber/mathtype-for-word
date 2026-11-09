# MathType for Word（中文整合版）

把 TeX 公式写成真正的 **MathType 7 Equation.DSMT4** 对象，在 Word DOCX / PowerPoint PPTX 里排版、编号、交叉引用，并就地编辑文档里已有的公式。全程走桌面版 Word/PowerPoint 的 COM 自动化，不生成 Word 自带 OMath、公式图片，也不用题注文字冒充公式。

本仓库是 [felimet/mathtype-for-word](https://github.com/felimet/mathtype-for-word) 的分支，保留上游 MIT 许可与署名。它在 **v1.4.3** 的基础上加入「就地编辑已有公式 + 中文学术排版」两套上游没有的能力；上游的批量渲染、纯文本公式扫描打标、MTEF 校验等能力已经同步进来。唯一的例外是 PowerShell 7 自动安装：本分支要求本机已装 `pwsh.exe`。

## 与上游的关系

- 同步基线：上游 **v1.4.3**（2026-11-08，`340498e`）。
- 本分支保留上游的 scan / mark、批量渲染与断点续跑、桥锁与孤儿 Office 清理、警告偏好 journal、行距/表格行高裁剪修复、MTEF 红字/菱形校验、已有编号计数修复。
- 本分支不包含上游的 PowerShell 7 自动安装、`tests/`、`evals/`、CI、插件市场打包；安装形态是单个自包含技能目录。

## 本分支新增（上游没有）

| 能力 | 说明 |
|---|---|
| 就地编辑已有公式 | `inventory` 列出每行公式的序数、布局与缓存 TeX；`edit` 按 `old_tex` 校验后原地重写，复用原定界符，支持删除与正文精确替换，校验失败不发布 |
| 公式预览 | `preview` 用 LibreOffice 把 DOCX 转成 PDF 与逐页 PNG，供目检排版 |
| 中文学术排版 | `其中，` 段零首行缩进、内联公式 `w:position=0` 与 `w:textAlignment=center`、禁止固定行距溢出；`normalize-word-inline-math-layout.py --check` 自动检查 |
| Texvc 预检 | 7.8 实测命令清单在 render 前拦截已知会变红字/丢符号的命令 |
| 环境探测增强 | MathType 从 PATH、注册表、安装根目录发现；Word 模板按需加载；用户目录可写性、Vector 样式、模板加载方式检查 |
| 模板安装 | `install-template` 把 MathType Commands 复制到 Word STARTUP，使模板常驻 |
| 单技能安装 | 桥脚本、编辑模块、MCP server、批量运行器都在 `mathtype-for-word\` 内，复制即可用 |

## 从上游 v1.4.3 合入的能力

| 能力 | 说明 |
|---|---|
| scan / mark | 从纯文本公式扫描候选并打标，生成可渲染的标记文档 |
| 批量渲染 | 超过 `batch_size`（默认 40）时按 chunk 用新的 Word 会话渲染，带检查点、可断点续跑、可后台运行并查询状态 |
| 桥锁与清理 | 并发编辑调用串行化；异常退出后清理孤儿 Word/PowerPoint；警告偏好写入 journal，下次运行恢复 |
| MTEF 校验 | 直接读取每个 `Equation.DSMT4` 的 MTEF 数据，红字（CHAR option bit 0x80）或菱形（U+FFFD）都会让 render 拒绝发布，并报出公式与命令 |
| 已有编号计数 | 往已经带 MathType 编号/引用的文档里追加公式不再被 validate 误判；结果报告 `existing_number_fields` / `existing_references` |
| 行距与行高修复 | 精确行距/表格行高低于公式时自动改为 At Least，并报告受影响的公式 |
| 渲染即校验 | render 发布前校验完整文档，成功结果自带 validation，省掉重复的 validate 调用 |

## 不包含的上游能力

- PowerShell 7 自动下载与安装（本分支要求 `pwsh.exe` 已存在，可用 `MATHTYPE_WORD_POWERSHELL` 指定绝对路径）。
- `tests/`、`evals/`、CI、`dist/` 发布包与 Codex/Claude 插件市场清单。

## 目录

| 路径 | 作用 |
|---|---|
| `mathtype-for-word/SKILL.md` | 技能契约：工作流、校验门、编号与引用策略、失败处理 |
| `mathtype-for-word/scripts/mathtype-word.ps1` | Word / PowerPoint 的 COM 桥 |
| `mathtype-for-word/scripts/mathtype-word-editing.ps1` | 发现、盘点、就地编辑、预览模块（被桥 dot-source） |
| `mathtype-for-word/scripts/mathtype_batch.py` | 分批、可续跑的渲染运行器 |
| `mathtype-for-word/scripts/mcp_server.py`、`run-mcp.ps1` | 本机 stdio MCP server（15 个工具） |
| `mathtype-for-word/scripts/invoke-bridge.ps1` | 解析同目录桥脚本的包装器 |
| `mathtype-for-word/scripts/normalize-word-inline-math-layout.py` | 内联公式排版检查与原地修正 |
| `mathtype-for-word/references/` | 学术公式排版、编号与引用、PPT 工作流、安装矩阵、排错、manifest 结构 |
| `scripts/install-skill.ps1` | 把技能安装到 Codex / 共享 agents / DSH 的技能根目录 |

桥脚本的动作：`probe`、`probe-pptx`、`configure-defaults`、`check-manifest`、`scan`、`mark`、`render`、`validate`、`edit`、`inventory`、`preview`、`install-template`、`update`、`render-pptx`、`validate-pptx`、`cleanup`。

## 前置条件

- Windows 10 或 11，Microsoft **Word / PowerPoint 桌面版**。Microsoft 365 的网页式 MathType Add-in 不行，本工具依赖桌面 OLE 与 Word 模板。
- **MathType for Windows**（上游以 7.11.1.462 开发验证，本分支在 7.8.2.441 上完成全部实测）。只装 Add-in 不够。
- PowerShell 7（`pwsh`），已安装；MCP server 另需 PATH 上有 Python 3。
- 可选：LibreOffice 与 Poppler（`pdftoppm`），用于 `preview` 生成 PDF/PNG。

动手前先跑 `probe`：只读，报告 MathType、Word 模板、用户目录可写性、Vector 样式等检查项；失败就停，不要硬跑。

## 安装

1. **作为技能安装（推荐）**

   ```powershell
   pwsh -File scripts\install-skill.ps1 -DefaultRoots
   ```

   默认写入 `%USERPROFILE%\.codex\skills\`、`%USERPROFILE%\.agents\skills\`、`%USERPROFILE%\.dsh\skills\`；也可用 `-SkillRoot` 指定任意技能根目录，用 `-Force` 覆盖已有安装。目标目录里的技能名固定为 `mathtype-for-word`。

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
- 本分支不自动安装 PowerShell 7；缺少 `pwsh.exe` 时桥和 MCP server 会直接报错。

## 版本与许可

- 本分支版本 **2.1.0**（2026-11-09），同步上游 v1.4.3，MIT 许可。
- 上游项目：<https://github.com/felimet/mathtype-for-word>
- 版权：上游 `LICENSE` 保留 Jia-Ming Zhou（Felimet）署名；本分支的修改同样以 MIT 发布。
