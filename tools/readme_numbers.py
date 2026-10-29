#!/usr/bin/env python3
"""README 数字重数器：一条命令出全部对外数字，杜绝详情页数字过期。

跑法：python3 tools/readme_numbers.py
每个数字都从代码/语料实数，不许手填。
"""
import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def mcp_tools() -> list[str]:
    src = (ROOT / "server/jszz/server.py").read_text("utf-8")
    # @mcp.tool() 装饰的下一个 def 即工具名
    return re.findall(r"@mcp\.tool\(\)\s*\ndef\s+(\w+)", src)


def main() -> None:
    py_lines = sum(len(p.read_text("utf-8").splitlines())
                   for p in (ROOT / "server/jszz").glob("*.py"))
    refs = sorted((ROOT / "skill/references").glob("*.md"))
    samples = sorted((ROOT / "examples/subjects").glob("*.md"))
    tools = mcp_tools()
    big = (ROOT / "server/jszz/pptx_render.py").read_text("utf-8")
    m = re.search(r'BIG = re.compile\(r"\^\((.*?)\)"\)', big, re.S)
    words = {w.strip() for w in re.split(r"[|]", m.group(1) if m else "", ) if len(w.strip()) > 1}
    tests = sorted((ROOT / "tests").glob("test*.py")) + [ROOT / "tests/selftest.py"]
    skill_zip = ROOT / "dist/jszz.skill"
    skill_files = len(zipfile.ZipFile(skill_zip).namelist()) if skill_zip.exists() else 0

    corpus = (ROOT / "docs/corpus-analysis.md").read_text("utf-8")
    full4 = ROOT / "analysis/full4/FINDINGS.md"
    deep = full4.read_text("utf-8") if full4.exists() else ""

    print(f"MCP 工具        {len(tools)}")
    print(f"  {', '.join(tools)}")
    print(f"排版规范        {len(refs)} 篇")
    print(f"样例            {len(samples)} 份")
    print(f"服务端代码      {py_lines} 行（server/jszz/*.py）")
    print(f"栏目词表        {len(words)} 词（pptx BIG）")
    print(f"测试            {len(tests)} 套")
    print(f"jszz.skill      {skill_files} 文件")
    print("--- 语料（docs/corpus-analysis.md）---")
    for line in corpus.splitlines()[:14]:
        if re.search(r"\d", line) and line.strip().startswith(("-", "|")):
            print(" ", line.strip())
    print("--- 深读（analysis/full4/FINDINGS.md）---")
    for line in deep.splitlines()[:4]:
        if line.strip():
            print(" ", line.strip())


if __name__ == "__main__":
    main()
