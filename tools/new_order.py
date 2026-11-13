#!/usr/bin/env python3
"""新建订单工作区：python3 tools/new_order.py "王老师_说课课件" """
import sys, datetime, pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent

def main():
    if len(sys.argv) < 2:
        sys.exit('用法：python3 tools/new_order.py "订单名"')
    name = sys.argv[1].strip().replace('/', '_')
    today = datetime.date.today().isoformat()
    d = ROOT / '订单' / f'{today}_{name}'
    if d.exists():
        sys.exit(f'已存在：{d}')
    for sub in ('材料', '工作', '成品/对比'):
        (d / sub).mkdir(parents=True)
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
    for tpl, out in (('理解卡.md', '理解卡.md'), ('问答记录.md', '问答记录.md')):
        text = (ROOT / '模板' / tpl).read_text(encoding='utf-8')
        text = text.replace('{订单名}', name).replace('{YYYY-MM-DD HH:MM}', now)
        (d / out).write_text(text, encoding='utf-8')
    (d / '材料' / '_转写.md').write_text(f'# 材料转写：{name}\n\n', encoding='utf-8')
    print(d)

if __name__ == '__main__':
    main()
