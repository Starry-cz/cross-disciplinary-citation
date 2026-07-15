# 跨学科参考文献与交叉引用

面向自然科学、社会科学、人文学科、医学、工程和计算机科学的 Codex Skill。它帮助选择正确的引文体系、核对多类型来源、生成一致的文内引文与文后条目，并为 DOCX 提供可点击 Word REF 交叉引用。

## 覆盖范围

- 体系：GB/T 7714、APA、Chicago（Author-Date / Notes-Bibliography）、MLA、IEEE、Vancouver/NLM、AMA、ACS、CSE，以及法学辖区格式。
- 来源：期刊、会议、专著、书章、学位论文、报告、预印本、数据集、软件、代码、标准、专利、网页、法规/判例、视听资料、个人通信和 AI 辅助内容。
- 文内形式：作者—年份、顺序编码、作者—页码、脚注/尾注、法学引证。
- Word：作者—年份精确映射为 REF 域；顺序编码交叉引用可配合 [ref-citation-embedder](https://github.com/hey123760/ref-citation-embedder)。

## 安装

```powershell
git clone https://github.com/Starry-cz/cross-disciplinary-citation.git "$HOME\.codex\skills\cross-disciplinary-citation"
```

刷新 Codex 后，技能显示为“跨学科参考文献与交叉引用”。

## 快速开始

1. 确认目标期刊、学校或会议的作者指南和指定版本。
2. 从 `assets/bibliography-record-template.yaml` 建立可验证的文献记录。
3. 阅读 `references/style-routing.md` 选择体系，并读取与来源类型对应的参考文件。
4. 生成正文和文后条目后，运行 `scripts/validate_skill.py .`。
5. 作者—年份 DOCX 交叉引用使用 `assets/author-date-map-template.txt` 和 `scripts/author_year_refcite.py`；编号制使用上游工具。

## 安全边界

- 不编造书目信息，不把搜索摘要或模型记忆当原始文献。
- 投稿指南优先；法学引证、机构政策和数据/软件引用须采用对应官方规则。
- 预印本、正式版、勘误和撤稿版本必须明确区分。
- 脚本只根据用户确认的精确映射写入 Word 域，并另存输出文件。

## 发布校验

```powershell
python scripts/validate_skill.py .
```

## 目录

```text
cross-disciplinary-citation/
├─ SKILL.md
├─ agents/openai.yaml
├─ references/
├─ assets/
├─ examples/
├─ scripts/
└─ README.md
```
