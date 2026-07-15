#!/usr/bin/env python3
"""校验技能发布结构、资源路由和 Python 脚本语法。"""

import re
import sys
from pathlib import Path


REQUIRED = (
    "SKILL.md",
    "agents/openai.yaml",
    "references/citation-patterns.md",
    "references/word-maintenance-workflow.md",
    "references/style-routing.md",
    "references/source-types.md",
    "references/digital-and-legal-sources.md",
    "references/citation-integrity.md",
    "assets/bibliography-record-template.yaml",
    "assets/author-date-map-template.txt",
    "examples/interdisciplinary-records.yaml",
    "scripts/author_year_refcite.py",
    "scripts/numeric_refcite.py",
    "scripts/reference_auto_numbering.py",
    "README.md",
)


def main():
    root = Path(sys.argv[1] if len(sys.argv) > 1 else ".").resolve()
    missing = [relative for relative in REQUIRED if not (root / relative).is_file()]
    if missing:
        raise ValueError("缺少发布文件：" + "、".join(missing))

    skill = (root / "SKILL.md").read_text(encoding="utf-8")
    if not re.match(r"^---\nname: cross-disciplinary-citation\ndescription: .+\n---", skill):
        raise ValueError("SKILL.md 前置元数据无效")

    # 仅检查 Markdown 中的本地相对链接，避免发布后出现失效资源路由。
    links = re.findall(r"\[[^]]+\]\(([^)]+)\)", skill)
    broken = [link for link in links if not link.startswith(("http://", "https://", "#")) and not (root / link).is_file()]
    if broken:
        raise ValueError("SKILL.md 存在失效本地链接：" + "、".join(broken))

    metadata = (root / "agents/openai.yaml").read_text(encoding="utf-8")
    for field in ("display_name:", "short_description:", "default_prompt:"):
        if field not in metadata:
            raise ValueError(f"agents/openai.yaml 缺少 {field}")
    if "$cross-disciplinary-citation" not in metadata:
        raise ValueError("default_prompt 必须显式调用 $cross-disciplinary-citation")

    # 编译脚本能快速发现发布前的语法错误。
    compile((root / "scripts/author_year_refcite.py").read_text(encoding="utf-8"), "author_year_refcite.py", "exec")
    compile((root / "scripts/numeric_refcite.py").read_text(encoding="utf-8"), "numeric_refcite.py", "exec")
    compile((root / "scripts/reference_auto_numbering.py").read_text(encoding="utf-8"), "reference_auto_numbering.py", "exec")
    print("技能发布结构、资源路由和脚本语法均通过校验。")


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError) as error:
        print(f"错误：{error}", file=sys.stderr)
        sys.exit(1)
