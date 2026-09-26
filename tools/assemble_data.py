#!/usr/bin/env python3
# tools/assemble_data.py —— raw + pairings + new_sentences → data.js
import json, os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOKS = [('b1', '成语练习（1-20）', 1, 20), ('b2', '成语练习（21-40）', 21, 40),
         ('b3', '成语练习（41-60）', 41, 60), ('b4', '成语练习（61-80）', 61, 80)]

def load(p):
    with open(os.path.join(ROOT, p), encoding='utf-8') as f:
        return json.load(f)

def main():
    pairings = load('tools/pairings.json')
    new_sents = {e['id']: e['sentences'] for e in load('tools/new_sentences.json')}
    entries = {}   # no -> {id, book, idiom, meaning, sentences}

    for bid, name, lo, hi in BOOKS:
        raw = load(f'tools/raw/{bid}.json')
        no_by_text = {i['idiom']: i['no'] for i in raw['idioms']}
        pair = pairings.get(bid)
        for sec in ('section1', 'section2'):
            for idx, e in enumerate(raw[sec]):
                ans = e['answer'] or (pair[sec][idx] if pair else None)
                assert ans in no_by_text, f'{bid} {sec} q{e["qno"]} 答案无效: {ans}'
                no = no_by_text[ans]
                it = entries.setdefault(no, {'id': no, 'book': bid, 'idiom': ans,
                                             'meaning': None, 'sentences': []})
                if sec == 'section1':
                    assert it['meaning'] is None, f'{bid} q{e["qno"]} 释义重复'
                    it['meaning'] = e['meaning']
                else:
                    it['sentences'].append({'text': e['sentence'], 'source': 'original'})

    assert sorted(entries) == list(range(1, 81)), '成语编号不完整'
    for no, it in entries.items():
        assert it['meaning'], f'#{no} 缺释义'
        assert len(it['sentences']) == 1, f'#{no} 原卷例句数异常'
        for s in new_sents[no]:
            it['sentences'].append({'text': s, 'source': 'new'})

    data = {'meta': {'title': '中一高级华文 成语复习', 'version': '3.0'},
            'books': [{'id': b, 'name': n} for b, n, _, _ in BOOKS],
            'idioms': [entries[no] for no in sorted(entries)]}

    with open(os.path.join(ROOT, 'data.js'), 'w', encoding='utf-8') as f:
        f.write('window.IDIOM_DATA = ')
        f.write(json.dumps(data, ensure_ascii=False, indent=2))
        f.write(';\n')
    print(f'data.js written: {len(data["idioms"])} idioms')

if __name__ == '__main__':
    main()
