# 本分支与上游的差异清单

- 上游基准：`940fc37`（v1.4.2 之后，2026-11-04），106 个受版本控制的文件，桥脚本 65 个函数，
  11 个 MCP 工具，105 个测试。
- 本分支基线：`441daf4`（2026-07-22），之后独立开发。仓库现有 22 个文件，桥脚本 42 个函数，
  编辑模块 25 个函数，12 个 MCP 工具，没有测试。

## 本分支有、上游没有

| 能力 | 落点 | 规模 |
|---|---|---|
| 就地编辑已有公式（inventory / edit） | `mathtype-word-editing.ps1` + 桥的 action 分派 | 编辑模块约 38 KB；含 `old_tex` 校验、定界符复用、删除、prose 精确替换、失败不发布 |
| 公式预览（preview） | 编辑模块，LibreOffice 转 PDF/PNG | 约 40 行，含转换器发现与输出等待 |
| 模板安装（install-template） | 编辑模块 | 把 MathType Commands 复制到 Word STARTUP |
| 环境探测增强 | 编辑模块 + 桥的 probe | MathType 可执行文件发现、模板运行时加载、`user_profile_writable`、`word_document_creation`、Vector 样式检查 |
| 中文学术排版规则 | `references/academic-equation-style.md`、`SKILL.md` | `其中，` 零缩进、`w:position=0`、`w:textAlignment=center`、At Least 行距 |
| 内联公式排版检查器 | `normalize-word-inline-math-layout.py` | 365 行，`--check` 只读，修复模式校验 OLE 字节不变 |
| Texvc 实测表 | `references/academic-equation-style.md`、`SKILL.md`、`troubleshooting.md` | 渲染成功 / 红字 / 黑菱形 / 静默丢符号四类，附替换表 |
| Texvc 预检与回读 | 桥的 `Get-UnsupportedTexCommand`、render/validate 的 `stored_tex` | render 前置拒绝、validate 回读对象内的 TeX 再判定 |

## 上游有、本分支没有

| 能力 | 风险 |
|---|---|
| `scan`（纯文本公式扫描）与 `mark`（打标） | 没有纯文本公式的自动定位，只能手写 manifest |
| 批量渲染、断点续跑、后台任务状态（`mathtype_batch.py`） | 单个 Word 会话转换几百次后会 RPC 报错，超过 40 个公式的长文档有失败风险 |
| 桥锁、孤儿 Word 清理、偏好 journal | 并发调用可能互相干扰；异常退出后残留隐藏 Word |
| 行距/表格行高裁剪修复（`exact_line_spacing`） | 精确行距下的公式会被裁切，只能事后用检查器发现 |
| PowerShell 7 自动安装（`ensure-pwsh.ps1`） | 要求用户自己装好 `pwsh.exe` |
| 插件市场打包、`tests/`、`evals/`、CI、`dist/` | 没有回归测试和发布产物，质量依赖手工验证 |
| PowerPoint 相关细化（剪贴板占用报错、实例复用、`height_points` 为空） | 少数边界情况下提示信息较少 |

## 移植成本判断

| 增量 | 能否独立提交 | 成本 |
|---|---|---|
| Texvc 实测表 + 中文排版规则（文档） | 能，已准备 `docs/mathtype-texvc-subset` | 低 |
| 内联公式检查器 + 测试 | 能，已准备 `feat/inline-math-layout-check`；依赖 lxml | 低 |
| Texvc 预检与回读 | 能，但不依赖编辑模块 | 中，需要接入 1.4.x 的 render/validate 流程 |
| 就地编辑 | 不能单文件搬运 | 高，需适配桥锁、偏好 journal、批量渲染的 chunk 状态 |
| 环境探测增强 | 与编辑模块耦合，可拆 | 中 |

## 已准备好的上游分支

- `docs/mathtype-texvc-subset`：`academic-equation-style.md` 的 Texvc 表与中文排版规则、
  `troubleshooting.md` 三行、`SKILL.md` 校验门与失败处理指向；dist 已重建。
- `feat/inline-math-layout-check`：检查器脚本、3 个单元测试、`SKILL.md` 工作流第 7 步与
  参考资料路由；dist 已重建。

两个分支都基于上游 `940fc37`，测试结果与 `main` 一致（仅本机 `python` 别名导致的 2 个环境错误）。
