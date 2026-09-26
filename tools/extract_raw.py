#!/usr/bin/env python3
# tools/extract_raw.py —— 解析 extracted/*.txt → tools/raw/b1..b4.json
import json, os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BOOKS = [
    ('b1', '成语练习（1-20）', 1, 20, True),
    ('b2', '成语练习（21-40）', 21, 40, True),
    ('b3', '成语练习（41-60）', 41, 60, False),
    ('b4', '成语练习（61-80）', 61, 80, False),
]

RE_NUM_INLINE = re.compile(r'^(\d+)[.、．]\s*(\S.*)$')   # "1.爱屋及乌" / "1.你成天这样…"
RE_NUM_ONLY = re.compile(r'^(\d+)[.、．]\s*$')            # "1." / "2．"
RE_NUM_ALONE = re.compile(r'^(\d+)$')                      # "41"（网格格式）
RE_LEADING_NUM = re.compile(r'^\d')

# 原卷文字勘误表（经主人确认 2026-09-26）：(book, section2 qno) -> (原文, 修正)
SENTENCE_FIXES = {
    ('b3', 16): ('这位球队', '这支球队'),   # 原卷量词语病
}


def read_lines(path):
    with open(path, encoding='utf-8') as f:
        return [ln.strip() for ln in f.read().splitlines() if ln.strip()]


def join_continuation(lines, i, text):
    """处理「城门失火，/殃及池鱼」跨行：text 以「，」结尾且下一行不以数字开头 → 拼接。"""
    while text.endswith('，') and i + 1 < len(lines) and not RE_LEADING_NUM.match(lines[i + 1]):
        i += 1
        text += lines[i]
    return text, i


def parse_idioms(lines):
    out, i = [], 0
    while i < len(lines):
        s = lines[i]
        m = RE_NUM_INLINE.match(s)
        if m:
            text, j = join_continuation(lines, i, m.group(2).strip())
            out.append((int(m.group(1)), text)); i = j + 1; continue
        if RE_NUM_ALONE.match(s) and i + 1 < len(lines):
            text, j = join_continuation(lines, i + 1, lines[i + 1])
            out.append((int(s), text)); i = j + 1; continue
        i += 1
    return out


def split_sections(lines):
    i1 = next(i for i, s in enumerate(lines) if s.startswith('一、'))
    i2 = next(i for i, s in enumerate(lines) if s.startswith('二、'))
    return lines[:i1], lines[i1 + 1:i2], lines[i2 + 1:]


def parse_section1(lines, idiom_set, expect_answers=False):
    entries, cur = [], None
    for s in lines:
        m = RE_NUM_ONLY.match(s)
        if m:
            cur = []; entries.append([int(m.group(1)), cur]); continue
        if cur is not None:
            cur.append(s)
    out, deferred, used = [], [], set()
    for qno, content in entries:
        answer = None
        if content:
            cand = content[0]
            if cand in idiom_set:
                answer, content = cand, content[1:]
            elif cand.endswith('，') and len(content) > 1 and cand + content[1] in idiom_set:
                answer, content = cand + content[1], content[2:]
            elif expect_answers and len(content) >= 2 and 2 <= len(cand) <= 12:
                # 答案卷里的错别字答案行（如 b2 q15「粗制滥做」）：先摘出，稍后近似配对
                deferred.append((qno, cand))
                content = content[1:]
        if answer:
            used.add(answer)
        out.append({'qno': qno, 'answer': answer, 'meaning': ''.join(content)})
    corrections = []
    if deferred:
        unused = idiom_set - used
        for qno, cand in deferred:
            best, best_score = None, 0
            for u in unused:
                score = len(set(cand) & set(u))
                if score > best_score:
                    best, best_score = u, score
            assert best and best_score >= 2, f'q{qno} 无法为错别字答案配对: {cand}'
            unused.discard(best)
            next(e for e in out if e['qno'] == qno)['answer'] = best
            corrections.append({'qno': qno, 'original': cand, 'corrected': best})
    return out, corrections


def parse_section2(lines, idiom_set):
    entries, unnumbered, cur = [], [], None
    i = 0
    while i < len(lines):
        s = lines[i]
        m_only = RE_NUM_ONLY.match(s)
        if m_only:
            cur = []; entries.append([int(m_only.group(1)), cur]); i += 1; continue
        m_inline = RE_NUM_INLINE.match(s)
        if m_inline:
            no, rest = int(m_inline.group(1)), m_inline.group(2).strip()
            cand, j = join_continuation(lines, i, rest)
            if cand in idiom_set:          # 第二大题前重复的选项表 → 跳过
                cur = None; i = j + 1; continue
            cur = [cand]; entries.append([no, cur]); i = j + 1; continue
        if cur is not None:
            cur.append(s)
        else:
            unnumbered.append(s)
        i += 1
    if unnumbered and not entries:          # 61-80：无题号的纯句子行
        entries = [[k + 1, [s]] for k, s in enumerate(unnumbered)]
    return entries


def normalize_blanks(s):
    s = re.sub(r'_{2,}|[\s　]{2,}', '____', s)
    return re.sub(r'(?:____)+', '____', s).strip()


def extract_answer(sentence, idiom_texts):
    for t in sorted(idiom_texts, key=len, reverse=True):
        m = re.search(r'[_\s]*' + re.escape(t) + r'[_\s]*', sentence)
        if m:
            return t, sentence[:m.start()] + '____' + sentence[m.end():]
    return None, sentence


def build(bid, name, lo, hi, has_answers):
    lines = read_lines(os.path.join(ROOT, 'extracted', name + '.txt'))
    head, s1_lines, s2_lines = split_sections(lines)
    idioms = sorted(parse_idioms(head), key=lambda x: x[0])   # 网格式表格按行主序读出，需按编号排序
    assert [n for n, _ in idioms] == list(range(lo, hi + 1)), f'{bid} 成语编号异常: {[n for n, _ in idioms]}'
    idiom_set = {t for _, t in idioms}
    assert len(idiom_set) == 20, f'{bid} 成语去重后不足 20'

    section1, corrections = parse_section1(s1_lines, idiom_set, expect_answers=has_answers)
    section2 = []
    for qno, frags in parse_section2(s2_lines, idiom_set):
        sentence = normalize_blanks(''.join(frags))
        answer = None
        if has_answers:
            answer, sentence = extract_answer(sentence, idiom_set)
            sentence = normalize_blanks(sentence)
        fix = SENTENCE_FIXES.get((bid, qno))
        if fix:
            assert fix[0] in sentence, f'{bid} section2 q{qno} 找不到勘误目标: {fix[0]}'
            sentence = sentence.replace(fix[0], fix[1])
        section2.append({'qno': qno, 'sentence': sentence, 'answer': answer})

    assert [e['qno'] for e in section1] == list(range(1, 21)), f'{bid} section1 题号异常'
    assert [e['qno'] for e in section2] == list(range(1, 21)), f'{bid} section2 题号异常'
    for e in section1:
        assert e['meaning'], f"{bid} section1 q{e['qno']} 释义为空"
    for e in section2:
        assert e['sentence'].count('____') == 1, f"{bid} section2 q{e['qno']}: {e['sentence']}"
    if has_answers:
        a1 = [e['answer'] for e in section1]
        a2 = [e['answer'] for e in section2]
        assert all(a1) and set(a1) == idiom_set, f'{bid} section1 答案缺失/不唯一'
        assert all(a2) and set(a2) == idiom_set, f'{bid} section2 答案缺失/不唯一'

    return {'book': {'id': bid, 'name': name, 'range': [lo, hi], 'hasAnswers': has_answers},
            'idioms': [{'no': n, 'idiom': t} for n, t in idioms],
            'section1': section1, 'section2': section2,
            'section1Corrections': corrections}


def main():
    out_dir = os.path.join(ROOT, 'tools', 'raw')
    os.makedirs(out_dir, exist_ok=True)
    for bid, name, lo, hi, has_answers in BOOKS:
        data = build(bid, name, lo, hi, has_answers)
        with open(os.path.join(out_dir, bid + '.json'), 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f'{bid}: idioms={len(data["idioms"])} s1={len(data["section1"])} s2={len(data["section2"])} OK')
    print('ALL OK')


if __name__ == '__main__':
    main()
