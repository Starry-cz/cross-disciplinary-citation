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

对于 IEEE、Vancouver、AMA、GB/T 7714 等 `[N]` 或上标编号体系，使用上游 [ref-citation-embedder](https://github.com/hey123760/ref-citation-embedder) 将已确认的编号标记转换为 REF 域。先决定每一处应引用哪条记录，再处理编号和链接；不要靠题名关键词自动猜测。

### 脚注/尾注与法学引证

Chicago Notes-Bibliography、OSCOLA、Bluebook、法院或本地法规格式通常需要页码、段落号、法域和缩略引注。先按目标法域或出版社指南生成注释；不将它们强制转换为编号文后条目。

## 硬性边界

- 不伪造或补全无法验证的书目信息；缺失字段必须标明缺失或向用户索取来源。
- 不把搜索摘要、二手转述或模型记忆当作已读原文；转引须明确，并以投稿指南为准。
- 不混用体系：同一文稿不要把作者—年份、编号、脚注格式拼接成未获许可的混合体系。
- 不重排、删改或替换用户文末条目，除非用户明确要求；Word 脚本默认另存输出并保留备份。
- 交付前运行 `python scripts/validate_skill.py .`，并以真实文献记录做代表性前向检查。
