# 中文作者年份引文与 Word 交叉引用

为中文社会科学、教育学和学位论文处理 GB/T 7714-2015 文后著录、著者—年份文内引文与 Word 可点击交叉引用的 Codex Skill。

它覆盖常见的 `（作者，年份）`、`(作者, 年份)`、`作者（年份）`、中英文双作者、`等` / `et al.`、多篇并列及同作者同年写法。图片所示的 `(景安磊等, 2024)` 可直接作为映射键。

## 功能边界

- 根据用户明确提供的“引文 = 参考文献编号”映射，在 DOCX 中将文内引文写为可点击的 Word REF 域。
- 为文末以 `[N]` 开头的对应文献加书签；不移动、重排或改写条目。
- 只做精确文本匹配，不根据题名、作者或主题猜测引用关系。

顺序编码制的 `[N]` 交叉引用和按首次出现顺序重排，请使用上游 [ref-citation-embedder](https://github.com/hey123760/ref-citation-embedder)。本项目在其 Word REF 域思路基础上，补足了著者—年份制的显示与映射工作流。

## 安装

```powershell
git clone https://github.com/Starry-cz/chinese-author-year-citation.git "$HOME\.codex\skills\chinese-author-year-citation"
```

刷新或重启 Codex 后，技能显示为“中文作者年份引文”。

## 使用

1. 让文末存在“参考文献”标题，且目标条目以唯一 `[N]` 开头。
2. 从 `assets/citation-map-template.txt` 复制映射文件，按 `精确引文=编号` 填写。例如：

   ```text
   景安磊等, 2024=1
   滕洋，2023=2
   Wang & Li, 2024=3
   ```

3. 先检查，再生成可点击引文：

   ```powershell
   python scripts/author_year_refcite.py 论文.docx --map-file citations.txt --check
   python scripts/author_year_refcite.py 论文.docx --map-file citations.txt --verify
   ```

4. 在 Word 中打开 `*_著者年份交叉引用.docx`，按 `Ctrl+A`、`F9` 更新域后，以 `Ctrl+单击` 验证跳转。

## 发布前校验

```powershell
python scripts/validate_skill.py .
```

该检查会确认核心文件、主技能资源链接、UI 元数据和交叉引用脚本语法。

## 目录

```text
chinese-author-year-citation/
├─ SKILL.md
├─ agents/openai.yaml
├─ references/
├─ assets/
├─ examples/
├─ scripts/
└─ README.md
```
