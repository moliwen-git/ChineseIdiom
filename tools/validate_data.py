#!/usr/bin/env python3
# tools/validate_data.py —— 校验 data.js 数据完整性（PASS 退出码 0 / FAIL 退出码 1）
import json, os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load_data():
    with open(os.path.join(ROOT, 'data.js'), encoding='utf-8') as f:
        text = f.read()
    m = re.match(r'^window\.IDIOM_DATA\s*=\s*([\s\S]*);\s*$', text)
    assert m, 'data.js 格式错误：应为 window.IDIOM_DATA = {…};'
    return json.loads(m.group(1))

def main():
    data = load_data()
    errors, warns = [], []
    idioms = data.get('idioms', [])
    if len(idioms) != 80: errors.append(f'成语数 {len(idioms)} != 80')
    if sorted(i.get('id') for i in idioms) != list(range(1, 81)):
        errors.append('id 不是 1-80 无缺号')
    per_book = {}
    for it in idioms:
        per_book[it.get('book')] = per_book.get(it.get('book'), 0) + 1
        if not isinstance(it.get('idiom'), str) or not it['idiom']:
            errors.append(f"#{it.get('id')} 成语缺失")
        if not it.get('meaning'): errors.append(f"#{it.get('id')} 释义为空")
        ss = it.get('sentences', [])
        srcs = [s.get('source') for s in ss]
        if len(ss) != 3 or srcs.count('original') != 1 or srcs.count('new') != 2:
            errors.append(f"#{it.get('id')} 例句应为 1 original + 2 new，实际 {srcs}")
        for s in ss:
            t = s.get('text', '')
            if t.count('____') != 1:
                errors.append(f"#{it.get('id')} 例句空格数异常: {t}")
            body = t.replace('____', '')
            if not (8 <= len(body) <= 50):
                warns.append(f"#{it.get('id')} 句长 {len(body)} 超出建议范围: {t}")
    texts = {}
    for it in idioms:
        for s in it.get('sentences', []):
            t = s.get('text', '')
            if t in texts:
                errors.append(f'例句重复: {t}（#{texts[t]} 与 #{it.get("id")}）')
            texts[t] = it.get('id')
    idiom_names = [it.get('idiom') for it in idioms]
    if len(idiom_names) != len(set(idiom_names)):
        errors.append('存在重名成语条目')
    for b, n in per_book.items():
        if n != 20: errors.append(f'{b} 成语数 {n} != 20')

    for w in warns: print('WARN:', w)
    if errors:
        for e in errors: print('FAIL:', e)
        sys.exit(1)
    print(f'PASS: {len(idioms)} idioms / {sum(len(i.get("sentences", [])) for i in idioms)} sentences，全部校验通过')

if __name__ == '__main__':
    main()
