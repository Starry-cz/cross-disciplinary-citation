---
name: chinese-author-year-citation
description: 为中文社会科学、教育学与学位论文规范 GB/T 7714-2015 参考文献、文内著者—年份引文和 Word 可点击交叉引用。适用于用户要求处理或校对“（作者，年份）”“作者（年份）”、顺序编码制 [N]、中英文混合、多文献并列引注，或将 DOCX 文内引文链接到文末参考文献时使用。
---

# 中文著者年份引文与交叉引用

## 先确认任务

1. 区分引文体系：著者—年份制、顺序编码制，或两者并存但分章节使用；不要混用。
2. 以用户给出的文后参考文献为准；不要凭题名或主题猜测来源。
3. 需要将 DOCX 变为可点击交叉引用时，保留原文件并另存输出。编号制可采用上游 `ref-citation-embedder` 的 `[N] → REF` 工作流；著者—年份制采用本技能脚本。

## 著者—年份制写法

- 中文括注：`（景安磊等，2024）`；若用户或期刊要求英文标点，可写 `(景安磊等, 2024)`。
- 叙述式：`景安磊等（2024）指出……`；不要重复写为“景安磊等（2024）……（景安磊等，2024）”。
- 两位作者：`（李木洲、孙艺源，2024）`；英文为 `(Wang & Li, 2024)`。
- 三位及以上作者：中文常用首位作者加“等”，英文常用首位作者加 `et al.`，但以目标期刊规定为准。
- 多篇并列：`（滕洋，2023；景安磊等，2024）`。每一项必须能在文后唯一定位。
- 同一作者同年多篇：在年份后加 `a`、`b`，例如 `（王明，2024a，2024b）`，并使文后条目同步。
- 转引、页码、机构作者、无日期和法规类条目的处理，读取 [references/citation-patterns.md](references/citation-patterns.md)。

## DOCX 交叉引用工作流

1. 检查文末存在“参考文献”标题，且每一条以唯一 `[N]` 编号开始；这是书签目标的唯一依据。
2. 创建 UTF-8 映射文件，每行 `正文中精确可见的引文=文后编号`。不要含外层括号。例如：

   ```text
   景安磊等, 2024=1
   滕洋，2023=2
   Wang & Li, 2024=3
   ```

   可复制 [assets/citation-map-template.txt](assets/citation-map-template.txt)，或参照 [examples/citations.example.txt](examples/citations.example.txt)。

3. 先用 `--check` 核对每个映射和出现次数，再执行写入：

   ```powershell
   python scripts/author_year_refcite.py 论文.docx --map-file citations.txt --check
   python scripts/author_year_refcite.py 论文.docx --map-file citations.txt --verify
   ```

4. 在 Word 打开输出文件后按 `Ctrl+A`、`F9` 更新域；按住 `Ctrl` 点击引文，确认能跳转到对应条目。

5. 发布或交付前运行 `python scripts/validate_skill.py .`，检查技能清单、核心文件和本地链接。

## 硬性规则

- 映射必须是精确文本匹配。找不到、同一引文映射到多个编号、参考文献编号重复或引文跨 Word run 时，停止并修复源文档或映射，不猜测、不自动替换。
- 不移动、重排或改写文后参考文献，只添加书签与 REF 域；需要按首次出现顺序重排时使用上游 `ref-citation-embedder`。
- 脚本只处理正文与参考文献标题之间的段落，跳过标题及文后条目。
- 写入前自动生成 `.bak.docx`；每次输出为 `*_著者年份交叉引用.docx`，不会覆盖原文件。
