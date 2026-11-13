#!/usr/bin/env python3
"""把项目的规则、技能、清单、模板、口吻、经验合成一份 Markdown，
给不能直接读仓库的 AI（网页版 ChatGPT / Grok / 豆包等）上传或粘贴使用。

用法：python3 tools/pack.py            → dist/接单助理_完整说明.md
      python3 tools/pack.py --lite     → 不含 playbook 全文（只做审阅和追问时用，体积小）
"""
import sys, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
LITE = '--lite' in sys.argv

ORDER = (
    ['AGENTS.md']
    + [f'skills/{s}/SKILL.md' for s in ('intake', 'ask', 'build', 'revise', 'retro')]
    + sorted(str(p.relative_to(ROOT)) for p in (ROOT / '清单').glob('*.md'))
    + sorted(str(p.relative_to(ROOT)) for p in (ROOT / '模板').glob('*.md'))
    + ['db/口吻/画像.md', 'db/口吻/范例.md']
    + ([] if LITE else ['db/经验/playbook.md'])
)

HEADER = '''# 接单助理 · 完整说明（打包版）

你将扮演"接单助理"。下面是全部规则与资料，各部分以 `=== 文件：路径 ===` 分隔。
你没有仓库文件系统：凡是说"写进某文件"的地方，改为在回复里输出该文件的完整新内容，
标明文件名，由主理人保存。凡是说"运行脚本"的地方，如果你不能运行代码，就直接说明。

'''

def main():
    out = [HEADER]
    for rel in ORDER:
        p = ROOT / rel
        if p.exists():
            out.append(f'\n\n=== 文件：{rel} ===\n\n' + p.read_text(encoding='utf-8'))
    dist = ROOT / 'dist'
    dist.mkdir(exist_ok=True)
    f = dist / ('接单助理_精简说明.md' if LITE else '接单助理_完整说明.md')
    f.write_text(''.join(out), encoding='utf-8')
    print(f, f'{f.stat().st_size // 1024} KB')

if __name__ == '__main__':
    main()
