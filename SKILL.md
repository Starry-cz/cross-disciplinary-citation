---
name: cross-disciplinary-citation
description: 为自然科学、社会科学、人文学科、医学、工程与计算机科学规范参考文献、文内引文、脚注/尾注和 Word 可点击交叉引用。适用于 APA、Chicago、MLA、IEEE、Vancouver/NLM、AMA、ACS、CSE、GB/T 7714 等体系，以及期刊论文、会议论文、专著、章节、学位论文、预印本、数据集、软件、代码仓库、标准、专利、技术报告、网页、法规与视听资料的著录、核对或 DOCX 引用链接任务。
---

# 跨学科参考文献与交叉引用

## 先确定规范

1. 确认目标期刊/学校、引文体系、版本与语言；投稿指南优先于通用规则。
2. 在 [references/style-routing.md](references/style-routing.md) 选择作者—年份、顺序编码、脚注/尾注或法学辖区引证。
3. 读取 [references/source-types.md](references/source-types.md)，按来源类型收集必要元数据；不补造作者、卷期、页码、DOI、版本或日期。
4. 涉及网页、预印本、软件、数据、代码、标准、专利、法规或判例时，读取 [references/digital-and-legal-sources.md](references/digital-and-legal-sources.md)。

## 规范与核对工作流

1. 从 [assets/bibliography-record-template.yaml](assets/bibliography-record-template.yaml) 复制记录模板，逐条填入可验证元数据。
2. 统一一个作品的版本：不要把预印本和正式发表版当成同一论断的两篇独立来源；无版本比较目的时优先正式版。
3. 核对作者顺序、题名、出处、年份、版本、卷(期)、页码/文章号、持久标识符与访问条件。DOI、软件版本、数据版本和代码 commit/tag 不可互换。
4. 依据所选体系生成正文引文和文后条目；文内每一项都必须唯一指向文后记录或脚注。
5. 按 [references/citation-integrity.md](references/citation-integrity.md) 审计：核实可追溯性、引文与论断匹配、转引、重版、撤稿/更正、个人通信与 AI 生成内容。

## Word 交叉引用

先按稿件所处阶段选择维护方式；详情见 [references/word-maintenance-workflow.md](references/word-maintenance-workflow.md)。

1. 写作和频繁修改阶段：优先保留 Zotero/EndNote、Word 自动编号或期刊模板，便于增删文献；不要过早冻结编号。
2. 定稿和投稿前：统一转换为文后正文字符编号 `[1]`、`[2]`、`[3]`，清除 Word 自动编号，再批量生成可点击 REF 域。
3. 需要用户后续在 Word 界面自行维护时：使用可见书签，如 `Ref_001`、`AYRef_001`；不要使用以下划线开头的隐藏书签。

### 作者—年份

仅处理文中精确可见的引文文本，例如 `(Smith et al., 2025)`、`（王明等，2024）`。文末对应条目必须以唯一 `[N]` 开头。

1. 从 [assets/author-date-map-template.txt](assets/author-date-map-template.txt) 创建映射，每行 `精确引文=文后编号`，不含外层括号。
2. 先检查，再写入：

   ```powershell
   python scripts/author_year_refcite.py 论文.docx --map-file citations.txt --check
   python scripts/author_year_refcite.py 论文.docx --map-file citations.txt --verify
   ```

3. 在 Word 中 `Ctrl+A`、`F9` 更新域，再用 `Ctrl+单击` 验证跳转。

### 顺序编码

对于 IEEE、Vancouver、AMA、GB/T 7714 等 `[N]` 或上标编号体系，先决定每一处应引用哪条记录，再处理编号和链接；不要靠题名关键词自动猜测。

1. 文后条目应以可选中的正文字符编号开头，例如 `[1] 作者. 题名...`；若用户文档混入 Word 自动编号，先清除自动编号。
2. 生成交叉引用时默认使用可见书签 `Ref_001`、`Ref_002`，这样用户可在 Word 的“交叉引用 -> 书签 -> 书签文字”中继续维护。
3. 正文中的顺序编码引用必须显示为黑色 9 磅右上角小标；范围和并列形式如 `[3,27]`、`[13-18]` 保留原显示符号。
4. 推荐使用本技能脚本处理 DOCX：

   ```powershell
   python scripts/numeric_refcite.py 论文.docx --check
   python scripts/numeric_refcite.py 论文.docx --normalize-reference-list --verify
   ```

5. `--check` 显示“Word 自动编号 0 条”时，现有 `[N]` 已是可选中的正文字符，不要再改成自动编号；它适合定稿交叉引用。
6. `--check` 显示 Word 自动编号大于 0 条时，必须加 `--normalize-reference-list`；没有该参数时脚本必须停止，不能静默改写文后列表。
7. 如果文后条目只有 Word 自动编号、没有真实 `[N]` 文本，必须显式加 `--insert-missing-reference-numbers` 才能按条目顺序补入正文字符编号。

### 写作阶段的 Word 自动编号

仅当用户明确要把文后列表改为 Word 自动编号时使用；该操作会删除段首真实 `[N]`，但可继续使用编号段落交叉引用。

```powershell
python scripts/reference_auto_numbering.py 论文.docx --verify
powershell -ExecutionPolicy Bypass -File scripts/word_auto_refcite.ps1 -InputPath 论文_自动编号.docx -OutputPath 论文_自动编号_交叉引用.docx
```

第二条命令要求 Windows 已安装 Microsoft Word，调用 Word 原生“编号项 → 段落编号”交叉引用接口，生成黑色 9 磅上标 REF 域，并保留 Ctrl+单击跳转。新增条目或重排后，先更新自动编号，再按 `Ctrl+A`、`F9` 更新正文域；若新增了全新条目，重新运行第二条命令以建立书签。

脚本只处理“参考文献”标题后的连续 `[1]`、`[2]` 条目；发现缺号或乱序必须报错，不得猜测重排。

### 脚注/尾注与法学引证

Chicago Notes-Bibliography、OSCOLA、Bluebook、法院或本地法规格式通常需要页码、段落号、法域和缩略引注。先按目标法域或出版社指南生成注释；不将它们强制转换为编号文后条目。

## 硬性边界

- 不伪造或补全无法验证的书目信息；缺失字段必须标明缺失或向用户索取来源。
- 不把搜索摘要、二手转述或模型记忆当作已读原文；转引须明确，并以投稿指南为准。
- 不混用体系：同一文稿不要把作者—年份、编号、脚注格式拼接成未获许可的混合体系。
- 不重排、删改或替换用户文末条目，除非用户明确要求；Word 脚本默认另存输出并保留备份。
- 交付前运行 `python scripts/validate_skill.py .`，并以真实文献记录做代表性前向检查。
