# Word 参考文献维护流程

该流程用于 DOCX 论文稿件，尤其是 GB/T 7714、IEEE、Vancouver、AMA 等顺序编码制。目标是在“写作便利”和“投稿格式稳定”之间分阶段处理，而不是在初稿阶段就把所有引用冻结。

## 三阶段策略

| 阶段 | 推荐做法 | 原因 |
| --- | --- | --- |
| 写作和频繁修改 | 使用 Zotero/EndNote、Word 自动编号或期刊模板 | 增删文献时编号自动更新，适合探索和大改 |
| 定稿和投稿前 | 转换为正文字符编号 `[1]`、`[2]`，清除 Word 自动编号 | 投稿版复制、检查、交叉引用和版式更稳定 |
| 需要用户后续手工维护 | 使用可见书签 `Ref_001`、`AYRef_001` | Word “交叉引用”窗口能找到书签，便于新增或修正 |

## 编号类型判断

- 能被鼠标全选选中的 `[15]` 是正文字符编号，可作为交叉引用目标。
- 全选正文时选不到、只显示在段落左侧的 `[1]` 是 Word 自动编号，不是正文文本。
- 若同一条文献同时出现左侧自动编号和正文里的 `[22]`，说明存在双编号。处理时清除自动编号，只保留正文字符编号。

## 顺序编码 DOCX 处理要求

1. 参考文献表每条记录必须有唯一编号。定稿阶段优先使用真实文本 `[N]`，不要依赖自动列表编号。
2. 正文引用生成 REF 域后必须设置为上标；常用中文期刊投稿稿件使用黑色 9 磅上标。
3. 默认生成可见书签，不使用 `_CitationRef1` 这类隐藏书签。可见书签格式为 `Ref_001`、`Ref_002`。
4. 文内 `[3,27]` 和 `[13-18]` 保留原符号；脚本只把可见数字转换为 REF 域，不擅自扩展范围。
5. 脚本检测到跨 Word run 的编号时必须报错，要求先在 Word 中统一该处字符格式，避免半个编号变成链接。

## 推荐命令

写作阶段：把现有的正文字符 `[N]` 转成 Word 自动编号：

```powershell
python scripts/reference_auto_numbering.py 论文.docx --verify
```

这会删除段首的真实 `[N]`，并在 Word 左侧生成选不中的自动编号。需要同时保留正文的可点击上标引用时，继续执行：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/word_auto_refcite.ps1 -InputPath 论文_自动编号.docx -OutputPath 论文_自动编号_交叉引用.docx
```

该模式要求 Windows 安装 Microsoft Word，并由 Word 原生接口生成 `REF Ref_001 \n \h`：`\n` 读取目标自动编号的编号值，`\h` 支持 Ctrl+单击跳转。添加、删除或移动文后条目后，先更新自动编号，再按 `Ctrl+A`、`F9` 更新正文 REF 域；若加入了新的文献条目，重新运行命令以为该条建立可见书签。

先只读检查当前编号形态：

```powershell
python scripts/numeric_refcite.py 论文.docx --check
```

输出中“Word 自动编号 0 条”表示 `[N]` 是正文字符；“Word 自动编号大于 0 条”表示仍有 Word 列表编号。脚本不会在未明确传入 `--normalize-reference-list` 时静默删除自动编号。

已存在正文字符编号，并需要清除自动编号：

```powershell
python scripts/numeric_refcite.py 论文.docx --normalize-reference-list --verify
```

只有 Word 自动编号、没有正文 `[N]` 文本时，明确允许按条目顺序补编号：

```powershell
python scripts/numeric_refcite.py 论文.docx --normalize-reference-list --insert-missing-reference-numbers --verify
```

作者—年份引文需要可见书签时：

```powershell
python scripts/author_year_refcite.py 论文.docx --map-file citations.txt --verify
```

## 用户手工维护方式

新增一条参考文献后：

1. 在文后条目开头写入真实文本编号，如 `[29]`。
2. 选中该编号，进入 `插入 -> 书签`，新建 `Ref_029`。
3. 在正文需要引用处进入 `引用 -> 交叉引用`。
4. 选择 `引用类型：书签`，`引用内容：书签文字`，插入 `Ref_029`。
5. 选中新插入的正文引用，设置为上标、小字号、黑色。

若新增文献导致全篇编号重排，优先回到文献管理器或脚本重新批处理，不建议逐个手工改域。
