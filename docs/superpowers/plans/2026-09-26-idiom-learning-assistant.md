# 成语学习小助手 · 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把 4 份 Word 考卷里的 80 个成语做成单文件离线网页工具「成语学习小助手.html」：翻转卡片复习视图 + 拖拽填空测试视图，题库独立存于 data.js。

**Architecture:** 纯静态单页应用（一个 HTML，内联样式脚本，顶部标签切换两视图）+ 外挂 `data.js` 题库。数据管线：docx 提取文本 → Python 解析成 raw JSON → cc 人工配对（41-80 答案）与新编例句 → 组装 `data.js`。练习记录存 localStorage，驱动「错题优先」抽题权重。

**Tech Stack:** 原生 HTML/CSS/JS（无框架、零网络依赖）、Python 3.13 标准库（数据管线）、Node 24 `node:test`（逻辑测试台，跑完即删）。

**Spec:** `docs/superpowers/specs/2026-09-26-idiom-review-tool-design.md`（v3）

## Global Constraints

- 唯一工具文件名：**`成语学习小助手.html`**；`<title>` 与页面顶部大标题均为「成语学习小助手」
- 整个文件夹拷到任何电脑**双击即用**：禁止外部 CDN/字体/网络请求；只允许 `<script src="data.js">` 一个外挂文件
- `data.js` 格式：首行 `window.IDIOM_DATA = ` + 标准 JSON + 结尾 `;`（UTF-8）
- 例句空格统一 `____`（4 个下划线），每句**恰好 1 个**
- 新编例句：新加坡中小学生生活场景（学校/CCA/小贩中心/组屋/地铁/邻里/节庆等）、南洋女中中一高级华文水平、语境唯一确定该成语、不与原卷例句重复、句长 10-45 字
- 抽题权重：错题（`wrongLast && correctStreak < 2`）→ 未考过（`asked === 0`）→ 其余按 `lastAskedAt` 升序；每轮 `min(10, 成语总数)` 题
- localStorage 键名固定：`idiom.progress`、`idiom.settings`；所有读取必须 try/catch，损坏时回退默认值
- 界面简体中文、浅色系专业风、移动端自适应（桌面最大宽度 1080px 居中）
- 逻辑代码必须写在 HTML 内联脚本的 `/* LOGIC-START */ … /* LOGIC-END */` 标记区内（Node 测试台靠标记提取）
- Python 脚本：文件读写显式 `encoding='utf-8'`；运行时加 `PYTHONIOENCODING=utf-8`（Windows 控制台是 cp1252，直接 print 中文会崩）
- TDD：测试台先跑一次**确认全红**再写实现；测试台跑完即删，**证据（测试输出摘要）写进 commit 信息**；`tools/` 下 extract_raw / assemble_data / validate_data 三个脚本是交付物，保留
- ⛔ 主人验收关卡（Task 2、3、8）：必须停下等主人确认后才能继续
- 每个任务结束提交 git，commit 信息结尾加：`Co-Authored-By: Claude Code <noreply@anthropic.com>`
- 工作目录：`C:\Users\MX-GMKTEC-X1\MyWork\ChineseIdiom`（git 仓库已建好，分支 `main`）

## 文件结构总览（任务完成后）

```
ChineseIdiom/
├── 成语学习小助手.html      ← Task 5 骨架+逻辑 / Task 6 复习视图 / Task 7 测试视图
├── data.js                  ← Task 4 组装生成（80 成语 × 3 例句）
├── README.md                ← Task 8
├── tools/
│   ├── extract_raw.py       ← Task 1 解析 extracted/*.txt → tools/raw/b1..b4.json（保留）
│   ├── raw/b1.json … b4.json← Task 1 产物（提交入库）
│   ├── pairings.json        ← Task 2 产物：41-80 答案配对（提交入库）
│   ├── new_sentences.json   ← Task 3 产物：160 句新编例句（提交入库）
│   ├── assemble_data.py     ← Task 4 raw+pairings+new_sentences → data.js（保留）
│   └── validate_data.py     ← Task 4 数据完整性校验（保留）
├── docs/
│   ├── idiom-pairing-review.md ← Task 2 产物：配对清单（供主人抽查）
│   ├── superpowers/specs/…     ← 设计文档（已完成）
│   └── superpowers/plans/…     ← 本计划
├── input/                   ← 原始考卷（不动）
└── extracted/               ← docx 提取文本（Task 1 的输入）
```

---

### Task 1: 数据提取脚本 `tools/extract_raw.py`

把 `extracted/*.txt`（4 份考卷的提取文本）解析成结构化 JSON。两种卷面格式：
- **1-20 / 21-40（答案卷）**：第一大题「题号行 → 答案成语行 → 释义行」；第二大题先重复一遍成语选项表（`21. 长年累月` 样式，需跳过），再「题号行 → 句子行」，答案成语内嵌在句子中（前后有下划线或多空格）。
- **41-60 / 61-80（空白卷）**：第一大题「题号行 → 释义行」（无答案）；第二大题 41-60 是「题号+句子同行」（`1.你成天这样____地混日子…`），61-80 是**无题号的 20 行句子**（按顺序即题号 1-20），空格是多个连续空格。
- 成语列表区两种格式：`1.爱屋及乌`（号+成语同行）或 网格 `41`↵`得过且过`（号行+成语行）。特例：`29. 城门失火，`↵`殃及池鱼` 跨行需拼接（成语以「，」结尾且下一行不以数字开头 → 拼接）。

**Files:**
- Create: `tools/extract_raw.py`
- Create（临时测试台，绿灯后删除）: `tools/test_extract.py`
- Output: `tools/raw/b1.json`、`b2.json`、`b3.json`、`b4.json`

**Interfaces:**
- Produces: `tools/raw/bN.json` 结构（Task 2/4 依赖）：
```json
{
  "book": { "id": "b1", "name": "成语练习（1-20）", "range": [1, 20], "hasAnswers": true },
  "idioms": [ { "no": 1, "idiom": "爱屋及乌" } ],
  "section1": [ { "qno": 1, "answer": "悲天悯人 或 null", "meaning": "悲天：哀叹时世…" } ],
  "section2": [ { "qno": 1, "sentence": "…____…", "answer": "爱屋及乌 或 null" } ]
}
```

- [x] **Step 1: 写失败测试 `tools/test_extract.py`**

```python
# tools/test_extract.py —— 临时测试台，跑绿后删除
import json, os, subprocess, sys, unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load(b):
    with open(os.path.join(ROOT, 'tools', 'raw', f'{b}.json'), encoding='utf-8') as f:
        return json.load(f)

class TestExtract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        env = dict(os.environ, PYTHONIOENCODING='utf-8')
        r = subprocess.run([sys.executable, os.path.join(ROOT, 'tools', 'extract_raw.py')],
                           capture_output=True, text=True, encoding='utf-8', env=env, cwd=ROOT)
        if r.returncode != 0:
            raise AssertionError('extract_raw.py 运行失败:\n' + (r.stderr or r.stdout))

    def test_all_books_shape(self):
        for b, lo in [('b1', 1), ('b2', 21), ('b3', 41), ('b4', 61)]:
            d = load(b)
            self.assertEqual([i['no'] for i in d['idioms']], list(range(lo, lo + 20)), b)
            self.assertEqual(len({i['idiom'] for i in d['idioms']}), 20, b)
            self.assertEqual([e['qno'] for e in d['section1']], list(range(1, 21)), b)
            self.assertEqual([e['qno'] for e in d['section2']], list(range(1, 21)), b)
            for e in d['section1']:
                self.assertTrue(e['meaning'], f"{b} 释义为空 q{e['qno']}")
            for e in d['section2']:
                self.assertEqual(e['sentence'].count('____'), 1, f"{b} q{e['qno']}: {e['sentence']}")

    def test_b1_answers(self):
        d = load('b1')
        self.assertEqual(d['section1'][0]['answer'], '悲天悯人')
        self.assertEqual(d['section2'][0]['answer'], '爱屋及乌')
        self.assertIn('她也爱看足球比赛了', d['section2'][0]['sentence'])
        answers1 = [e['answer'] for e in d['section1']]
        answers2 = [e['answer'] for e in d['section2']]
        self.assertEqual(set(answers1), {i['idiom'] for i in d['idioms']})
        self.assertEqual(set(answers2), {i['idiom'] for i in d['idioms']})

    def test_b2_joined_idiom(self):
        d = load('b2')
        self.assertIn('城门失火，殃及池鱼', {i['idiom'] for i in d['idioms']})
        answers1 = [e['answer'] for e in d['section1']]
        self.assertEqual(set(answers1), {i['idiom'] for i in d['idioms']})

    def test_b3_b4_no_answers(self):
        for b in ('b3', 'b4'):
            d = load(b)
            self.assertTrue(all(e['answer'] is None for e in d['section1']), b)
            self.assertTrue(all(e['answer'] is None for e in d['section2']), b)
        self.assertEqual(load('b4')['section1'][0]['meaning'], '比喻基础深厚，不容易动摇。')
        self.assertIn('的做法，太冒险了', load('b4')['section2'][0]['sentence'])

if __name__ == '__main__':
    unittest.main(verbosity=2)
```

- [x] **Step 2: 跑测试确认全红**

Run: `cd C:\Users\MX-GMKTEC-X1\MyWork\ChineseIdiom && PYTHONIOENCODING=utf-8 python tools/test_extract.py`
Expected: ERROR/FAIL（`extract_raw.py` 不存在，subprocess 返回非 0）

- [x] **Step 3: 实现 `tools/extract_raw.py`**

```python
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


def parse_section1(lines, idiom_set):
    entries, cur = [], None
    for s in lines:
        m = RE_NUM_ONLY.match(s)
        if m:
            cur = []; entries.append([int(m.group(1)), cur]); continue
        if cur is not None:
            cur.append(s)
    out = []
    for qno, content in entries:
        answer = None
        if content:
            cand = content[0]
            if cand in idiom_set:
                answer, content = cand, content[1:]
            elif cand.endswith('，') and len(content) > 1 and cand + content[1] in idiom_set:
                answer, content = cand + content[1], content[2:]
        out.append({'qno': qno, 'answer': answer, 'meaning': ''.join(content)})
    return out


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
    s = re.sub(r'_{2,}|[\s\u3000]{2,}', '____', s)
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
    idioms = parse_idioms(head)
    assert [n for n, _ in idioms] == list(range(lo, hi + 1)), f'{bid} 成语编号异常: {[n for n, _ in idioms]}'
    idiom_set = {t for _, t in idioms}
    assert len(idiom_set) == 20, f'{bid} 成语去重后不足 20'

    section1 = parse_section1(s1_lines, idiom_set)
    section2 = []
    for qno, frags in parse_section2(s2_lines, idiom_set):
        sentence = normalize_blanks(''.join(frags))
        answer = None
        if has_answers:
            answer, sentence = extract_answer(sentence, idiom_set)
            sentence = normalize_blanks(sentence)
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
            'section1': section1, 'section2': section2}


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
```

- [x] **Step 4: 跑测试确认全绿**

Run: `cd C:\Users\MX-GMKTEC-X1\MyWork\ChineseIdiom && PYTHONIOENCODING=utf-8 python tools/test_extract.py`
Expected: 4 个测试全部 OK。若断言失败，打印对应 raw JSON 检查解析分支（常见：61-80 无题号句子、城门失火跨行拼接）。

- [x] **Step 5: 删除测试台，提交（证据写进 commit 信息）**

```bash
rm tools/test_extract.py
git add tools/extract_raw.py tools/raw/
git commit -m "feat(tools): docx 提取文本解析为结构化 raw JSON（4 册 × 20 成语）

测试台 tools/test_extract.py 全绿后删除：
- test_all_books_shape OK（4 册编号/题号/空格数断言）
- test_b1_answers OK（悲天悯人/爱屋及乌抽查 + 答案唯一性）
- test_b2_joined_idiom OK（城门失火，殃及池鱼跨行拼接）
- test_b3_b4_no_answers OK（空白卷无答案 + 语义抽查）

Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 2: 41-80 答案配对（⛔ 含主人抽查关卡）

41-60、61-80 两册原卷无答案。每册 20 个成语 ↔ 20 条释义 ↔ 20 道例句，由执行者（cc）做语义一对一配对。**配对不是写代码，是内容工作**：逐条读释义/例句，从该册 20 个成语里选唯一匹配项，并为每条写一句配对理由。

**Files:**
- Create: `tools/pairings.json`
- Create: `docs/idiom-pairing-review.md`
- Create（临时测试台，绿灯后删除）: `tools/test_pairings.py`

**Interfaces:**
- Consumes: `tools/raw/b3.json`、`tools/raw/b4.json`（Task 1 产物）
- Produces: `tools/pairings.json`（Task 4 依赖），结构：
```json
{
  "b3": { "section1": ["按 qno1-20 顺序的 20 个成语"], "section2": ["同前"] },
  "b4": { "section1": ["…"], "section2": ["…"] }
}
```

- [ ] **Step 1: 写失败测试 `tools/test_pairings.py`**

```python
# tools/test_pairings.py —— 临时测试台，跑绿后删除
import json, os, unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load(p):
    with open(os.path.join(ROOT, p), encoding='utf-8') as f:
        return json.load(f)

class TestPairings(unittest.TestCase):
    def test_permutation(self):
        pairings = load('tools/pairings.json')
        for b in ('b3', 'b4'):
            raw = load(f'tools/raw/{b}.json')
            idiom_set = {i['idiom'] for i in raw['idioms']}
            for sec in ('section1', 'section2'):
                arr = pairings[b][sec]
                self.assertEqual(len(arr), 20, f'{b}.{sec} 长度')
                self.assertEqual(set(arr), idiom_set, f'{b}.{sec} 必须是该册 20 个成语的一一配对')

if __name__ == '__main__':
    unittest.main(verbosity=2)
```

- [ ] **Step 2: 跑测试确认全红**

Run: `PYTHONIOENCODING=utf-8 python tools/test_pairings.py`
Expected: FAIL（`tools/pairings.json` 不存在，FileNotFoundError）

- [ ] **Step 3: 逐条语义配对，写出 `tools/pairings.json`**

做法：读 `tools/raw/b3.json`（61-80 同理读 b4），对 section1 的每条释义、section2 的每条例句，从该册 `idioms` 的 20 个成语中选出唯一语义匹配项。示例（b4）：
- 释义「比喻基础深厚，不容易动摇。」→ `根深蒂固`
- 例句「他这种____的做法，太冒险了。」→ `孤注一掷`
- 例句「我们做事应该____，不能鬼鬼祟祟。」→ `光明正大`

要求：每个成语在 section1、section2 中各恰好用一次（测试保证）；拿不准的先放最可能的，并在审查文档里标 ⚠️ 请主人重点看。

- [ ] **Step 4: 写 `docs/idiom-pairing-review.md` 审查清单**

格式（b3、b4 各两张表）：

```markdown
# 41-80 答案配对审查清单

> 原卷这两册未附答案，以下是 cc 的配对结果，请主人抽查（尤其带 ⚠️ 的条目）。

## 成语练习（41-60）· 第一大题（释义 → 成语）

| 题号 | 释义（原文摘录） | 配对成语 | 理由 |
|---|---|---|---|
| 1 | 独:唯独;善:搞好…只顾自己,不管别人 | 独善其身 | 释义即该成语字面拆解 |
…

## 成语练习（41-60）· 第二大题（例句 → 成语）

| 题号 | 例句（挖空） | 配对成语 | 理由 |
…
```

- [ ] **Step 5: 跑测试确认全绿**

Run: `PYTHONIOENCODING=utf-8 python tools/test_pairings.py`
Expected: PASS（4 组均为完整一一配对）

- [ ] **Step 6: ⛔ 主人抽查关卡**

把 `docs/idiom-pairing-review.md` 交给主人抽查（提醒可用 `cc :md` 阅读）。**等主人确认配对无误（或按主人指正修改后重跑 Step 5）才能继续**。

- [ ] **Step 7: 删除测试台，提交**

```bash
rm tools/test_pairings.py
git add tools/pairings.json docs/idiom-pairing-review.md
git commit -m "feat(data): 41-80 两册答案语义配对 + 主人抽查清单

测试台 tools/test_pairings.py 全绿后删除：
- test_permutation OK（b3/b4 × section1/section2 均为 20 成语一一配对）
主人抽查：已于 <日期> 确认 / 修正条目：<列出>

Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 3: 新编例句 160 句（⛔ 含主人抽读关卡）

给 80 个成语每个再配 **2 句 cc 新编例句**（原卷例句已有 1 句）。这是内容创作任务，按册分 4 批做（每批 20 个成语 40 句），保证质量。

**Files:**
- Create: `tools/new_sentences.json`
- Create（临时测试台，绿灯后删除）: `tools/test_sentences.py`

**Interfaces:**
- Consumes: `tools/raw/b1..b4.json`（对照原卷例句防重复）
- Produces: `tools/new_sentences.json`（Task 4 依赖），结构：
```json
[ { "id": 1, "sentences": ["新例句1（含一个____）", "新例句2（含一个____）"] } ]
```
80 个条目，`id` 为成语全局编号 1-80，每条恰好 2 句。

**内容要求（Global Constraints 的展开，写作时逐条对照）：**
- 场景库：学校（考试/班会/图书馆/值日）、CCA 活动（篮球/舞蹈/童军）、小贩中心、组屋邻里、地铁/巴士、公园、节庆（农历新年/中秋/国庆/屠妖节）、家庭（做家务/照顾弟妹）、同学交往（借笔记/合作/矛盾）等；
- 语境必须唯一锁定该成语：写完自问「换成同册另一个成语还通吗？」通就重写；
- 句长 10-45 字（不含 `____`），复杂度参照原卷：一个情境从句 + 一个成语位；
- 同一成语的 2 句新例句场景不能雷同，也不得与原卷例句场景重复；
- 用词符合中一学生水平（不用生僻字词、不用网络流行语）；
- 标点用全角（，。！？），`____` 前后不加空格。

- [ ] **Step 1: 写失败测试 `tools/test_sentences.py`**

```python
# tools/test_sentences.py —— 临时测试台，跑绿后删除
import json, os, re, unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def load(p):
    with open(os.path.join(ROOT, p), encoding='utf-8') as f:
        return json.load(f)

class TestNewSentences(unittest.TestCase):
    def setUp(self):
        self.data = load('tools/new_sentences.json')
        self.originals = {}   # id -> 原卷例句
        for b in ('b1', 'b2', 'b3', 'b4'):
            raw = load(f'tools/raw/{b}.json')
            no_by_text = {i['idiom']: i['no'] for i in raw['idioms']}
            for e in raw['section2']:
                ans = e['answer']
                if ans:
                    self.originals[no_by_text[ans]] = e['sentence']

    def test_structure(self):
        self.assertEqual(len(self.data), 80)
        self.assertEqual(sorted(e['id'] for e in self.data), list(range(1, 81)))
        for e in self.data:
            self.assertEqual(len(e['sentences']), 2, f"id={e['id']} 应为 2 句")

    def test_blank_and_length(self):
        for e in self.data:
            for s in e['sentences']:
                self.assertEqual(s.count('____'), 1, f"id={e['id']}: {s}")
                body = s.replace('____', '')
                self.assertTrue(10 <= len(body) <= 45, f"id={e['id']} 句长 {len(body)}: {s}")
                self.assertNotRegex(s, r'____\s+|\s+____', f"id={e['id']} 空格前后不应有空格: {s}")

    def test_no_duplicates(self):
        all_new = [s for e in self.data for s in e['sentences']]
        self.assertEqual(len(all_new), len(set(all_new)), '新例句内部有重复')
        for e in self.data:
            orig = self.originals.get(e['id'])
            if orig:
                for s in e['sentences']:
                    self.assertNotEqual(s, orig, f"id={e['id']} 与原卷例句重复")

if __name__ == '__main__':
    unittest.main(verbosity=2)
```

注意：`test_no_duplicates` 里原卷例句只有 b1/b2 有答案（b3/b4 配对在 Task 2 完成后才有 `pairings.json`）。本测试台只对照 b1/b2 的原句做防重，b3/b4 的原句防重由 Task 4 的 `validate_data.py` 兜底（组装后全量对照）。

- [ ] **Step 2: 跑测试确认全红**

Run: `PYTHONIOENCODING=utf-8 python tools/test_sentences.py`
Expected: FAIL（`tools/new_sentences.json` 不存在）

- [ ] **Step 3: 分 4 批创作 160 句，写入 `tools/new_sentences.json`**

每批流程：读该册 `tools/raw/bN.json` 的 20 个成语 + 原卷例句 → 按「内容要求」写 40 句 → 追加进 JSON。示例（id=1 爱屋及乌，原卷句是「男朋友是足球队员」场景，新句须避开恋爱/足球场景）：

```json
{ "id": 1, "sentences": [
  "弟弟本来不喜欢下棋，但因为____，常陪爷爷到组屋楼下棋社看人下棋。",
  "美玲加入摄影学会后，____，连逛街时都爱抬头拍组屋窗外的天空。"
] }
```

（第一句答案是爱屋及乌？——注意：这是**反例**，「因为____」语义不通。执行时写成：「弟弟本来不喜欢下棋，但____，因为爷爷爱下棋，他也常陪爷爷到楼下棋社。」创作时以「语境唯一锁定」为准，此处仅示意格式。）

- [ ] **Step 4: 跑测试确认全绿**

Run: `PYTHONIOENCODING=utf-8 python tools/test_sentences.py`
Expected: 3 个测试全部 OK

- [ ] **Step 5: ⛔ 主人抽读关卡**

从 160 句中随机抽 20 句 + 每册各 2 个成语的完整 3 句对照，贴给主人抽读，确认水平与场景贴合。**等主人认可（或按主人意见修改后重跑 Step 4）才能继续**。

- [ ] **Step 6: 删除测试台，提交**

```bash
rm tools/test_sentences.py
git add tools/new_sentences.json
git commit -m "feat(data): 80 成语 × 2 句新编例句（新加坡中学生生活场景）

测试台 tools/test_sentences.py 全绿后删除：
- test_structure OK（80 条 × 2 句，id 1-80 无缺）
- test_blank_and_length OK（每句恰 1 个____，句长 10-45）
- test_no_duplicates OK（新句互不重复、不与 b1/b2 原句重复）
主人抽读：已于 <日期> 确认 / 修改条目：<列出>

Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 4: 组装 `data.js` + 数据校验脚本

**Files:**
- Create: `tools/assemble_data.py`（保留）
- Create: `tools/validate_data.py`（保留，交付物）
- Output: `data.js`（根目录）

**Interfaces:**
- Consumes: `tools/raw/b1..b4.json`、`tools/pairings.json`、`tools/new_sentences.json`
- Produces: `data.js` —— `window.IDIOM_DATA = ` + JSON（spec §3.3 schema：`meta`/`books`/`idioms[{id,book,idiom,meaning,sentences[3]}]`）；`validate_data.py` 退出码 0=PASS、1=FAIL

- [ ] **Step 1: 写 `tools/assemble_data.py`**

```python
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
```

- [ ] **Step 2: 写 `tools/validate_data.py`（spec §8.1 全量校验，保留为交付物）**

```python
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
    texts_seen = {}
    for it in idioms:
        per_book[it.get('book')] = per_book.get(it.get('book'), 0) + 1
        if it.get('idiom') in texts_seen:
            errors.append(f"成语重复: {it['idiom']}")
        texts_seen[it.get('idiom')] = it.get('id')
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
    all_texts = [s['text'] for it in idioms for s in it.get('sentences', [])]
    if len(all_texts) != len(set(all_texts)):
        errors.append('240 句中存在完全重复的句子')
    for b, n in per_book.items():
        if n != 20: errors.append(f'{b} 成语数 {n} != 20')

    for w in warns: print('WARN:', w)
    if errors:
        for e in errors: print('FAIL:', e)
        sys.exit(1)
    print(f'PASS: {len(idioms)} idioms / {len(all_texts)} sentences，全部校验通过')

if __name__ == '__main__':
    main()
```

- [ ] **Step 3: 运行组装 + 校验**

```bash
cd C:\Users\MX-GMKTEC-X1\MyWork\ChineseIdiom
PYTHONIOENCODING=utf-8 python tools/assemble_data.py
PYTHONIOENCODING=utf-8 python tools/validate_data.py
```
Expected: `data.js written: 80 idioms` + `PASS: 80 idioms / 240 sentences，全部校验通过`（WARN 允许存在，逐条人工看过即可）

- [ ] **Step 4: 抽查 data.js 内容**

打开 `data.js` 头尾：首行 `window.IDIOM_DATA = {`、结尾 `};`、`#1 爱屋及乌` 有 3 句（1 original 来自考卷 + 2 new）、`#80 汗牛充栋` 同样齐全。

- [ ] **Step 5: Commit**

```bash
git add tools/assemble_data.py tools/validate_data.py data.js
git commit -m "feat(data): 组装 data.js 题库（80 成语 × 3 例句 = 240 句）

- assemble_data.py：raw + pairings + new_sentences 合并，断言释义/例句齐全
- validate_data.py：PASS（80 idioms / 240 sentences，spec §8.1 全量校验）

Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 5: HTML 骨架 + 内联核心逻辑（TDD，Node 测试台）

先建 `成语学习小助手.html` 骨架（含标签切换的双视图容器 + 完整 CSS），核心纯逻辑写在 `/* LOGIC-START */ … /* LOGIC-END */` 标记区内（UMD 风格：浏览器挂 `window.IdiomApp`，Node 测试台提取后当模块加载）。本任务只做逻辑与骨架，两个视图的内容渲染在 Task 6/7。

**Files:**
- Create: `成语学习小助手.html`
- Create（临时测试台，绿灯后删除）: `tools/logic.test.mjs`

**Interfaces:**
- Consumes: `data.js`（Task 4 产物）
- Produces: `window.IdiomApp`（Task 6/7 依赖，签名如下，全部纯函数除 makeStorage）：
  - `newProgress() → { byIdiom: {} }`
  - `makeStorage() → { available: boolean, backend: {getItem,setItem,removeItem} }`（localStorage 不可用时降级内存 Map）
  - `loadProgress(storage) → progress`（损坏/缺失 → `newProgress()`）
  - `saveProgress(progress, storage) → boolean`
  - `loadSettings(storage) → object` / `saveSettings(settings, storage) → boolean`（键 `idiom.settings`）
  - `shuffle(arr, rng = Math.random) → 新数组`（Fisher-Yates）
  - `mulberry32(seed) → rng 函数`（测试用确定性随机源）
  - `pickIdioms(idioms, progress, count = 10, rng = Math.random) → 成语对象数组`（权重：错题 → 未考过 → lastAskedAt 升序；返回 `min(count, idioms.length)` 个不重复项）
  - `buildRound(picked, rng = Math.random) → [{ slot, idiomId, sentenceIndex, text, answer, meaning }]`
  - `gradeRound(round, userAnswers) → { results: [{ slot, idiomId, correct, userAnswer, correctAnswer, meaning, text }], score, total }`（`userAnswers` 是按 slot 下标的成语文本数组）
  - `updateProgress(progress, round, results, now = Date.now()) → progress`（对→streak+1 且 streak≥2 时 wrongLast=false；错→streak=0、wrongLast=true；asked+1；lastAskedAt=now）
  - `validateData(data) → { ok: boolean, errors: string[] }`

- [ ] **Step 1: 写失败测试 `tools/logic.test.mjs`**

```js
// tools/logic.test.mjs —— 临时测试台，跑绿后删除
import test from 'node:test';
import assert from 'node:assert';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const ROOT = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const html = readFileSync(path.join(ROOT, '成语学习小助手.html'), 'utf8');
const m = html.match(/\/\* LOGIC-START \*\/([\s\S]*?)\/\* LOGIC-END \*\//);
assert.ok(m, 'HTML 缺少 LOGIC-START/LOGIC-END 标记区');
const module_ = { exports: {} };
new Function('module', 'exports', m[1])(module_, module_.exports);
const App = module_.exports;

const IDS = [...Array(80)].map((_, i) => ({
  id: i + 1, book: 'b' + (Math.floor(i / 20) + 1), idiom: `成语${i + 1}`, meaning: `释义${i + 1}`,
  sentences: [0, 1, 2].map(k => ({ text: `句${i + 1}-${k}____。`, source: k === 0 ? 'original' : 'new' })),
}));

test('shuffle 保持元素集合与长度', () => {
  const rng = App.mulberry32(42);
  const out = App.shuffle([1, 2, 3, 4, 5], rng);
  assert.deepStrictEqual(out.slice().sort(), [1, 2, 3, 4, 5]);
});

test('pickIdioms 全新进度：返回 10 个不重复', () => {
  const picked = App.pickIdioms(IDS, App.newProgress(), 10, App.mulberry32(1));
  assert.strictEqual(picked.length, 10);
  assert.strictEqual(new Set(picked.map(p => p.id)).size, 10);
});

test('pickIdioms 成语不足 10 个时返回全部', () => {
  assert.strictEqual(App.pickIdioms(IDS.slice(0, 5), App.newProgress(), 10).length, 5);
});

test('pickIdioms 错题优先', () => {
  const p = App.newProgress();
  p.byIdiom[7] = { asked: 3, correct: 1, correctStreak: 0, wrongLast: true, lastAskedAt: 100 };
  for (const id of [1, 2, 3]) p.byIdiom[id] = { asked: 1, correct: 1, correctStreak: 2, wrongLast: false, lastAskedAt: id };
  const picked = App.pickIdioms(IDS, p, 10, App.mulberry32(7));
  assert.ok(picked.some(x => x.id === 7), '错题 #7 必须入选');
  for (const id of [1, 2, 3]) assert.ok(!picked.some(x => x.id === id), '已掌握且刚考过的不应入选');
});

test('pickIdioms 连对 2 次不再享受错题优先', () => {
  const p = App.newProgress();
  p.byIdiom[5] = { asked: 4, correct: 3, correctStreak: 2, wrongLast: true, lastAskedAt: 999 };
  const picked = App.pickIdioms(IDS.slice(0, 10), p, 3, App.mulberry32(3));
  assert.ok(!picked.some(x => x.id === 5) || picked.every(x => x.id === 5 || p.byIdiom[x.id] === undefined || p.byIdiom[x.id].asked === 0),
    'streak>=2 时应与未考成语同权（本例 9 个未考，5 号不应挤进前 3）');
});

test('buildRound 结构与随机句', () => {
  const picked = IDS.slice(0, 10);
  const round = App.buildRound(picked, App.mulberry32(5));
  assert.strictEqual(round.length, 10);
  round.forEach((q, slot) => {
    assert.strictEqual(q.slot, slot);
    assert.ok(q.sentenceIndex >= 0 && q.sentenceIndex <= 2);
    assert.strictEqual(q.text, picked[slot].sentences[q.sentenceIndex].text);
    assert.strictEqual(q.answer, picked[slot].idiom);
  });
});

test('gradeRound 判分', () => {
  const round = App.buildRound(IDS.slice(0, 3), App.mulberry32(9));
  const answers = [round[0].answer, '成语999', null];
  const { results, score, total } = App.gradeRound(round, answers);
  assert.strictEqual(total, 3);
  assert.strictEqual(score, 1);
  assert.strictEqual(results[0].correct, true);
  assert.strictEqual(results[1].correct, false);
  assert.strictEqual(results[1].correctAnswer, '成语2');
  assert.strictEqual(results[2].userAnswer, null);
});

test('updateProgress 连对计数与错题标记', () => {
  const p = App.newProgress();
  const round = App.buildRound(IDS.slice(0, 2), App.mulberry32(2));
  let { results } = App.gradeRound(round, [null, round[1].answer]);
  App.updateProgress(p, round, results, 1000);
  assert.strictEqual(p.byIdiom[1].wrongLast, true);
  assert.strictEqual(p.byIdiom[1].correctStreak, 0);
  assert.strictEqual(p.byIdiom[2].correctStreak, 1);
  results = App.gradeRound(round, [round[0].answer, round[1].answer]).results;
  App.updateProgress(p, round, results, 2000);
  assert.strictEqual(p.byIdiom[1].correctStreak, 1);
  assert.strictEqual(p.byIdiom[1].wrongLast, true, 'streak=1 还未掌握');
  assert.strictEqual(p.byIdiom[2].correctStreak, 2);
  results = App.gradeRound(round, [round[0].answer, round[1].answer]).results;
  App.updateProgress(p, round, results, 3000);
  assert.strictEqual(p.byIdiom[1].wrongLast, false, '连对 2 次移出错题优先');
  assert.strictEqual(p.byIdiom[1].asked, 3);
  assert.strictEqual(p.byIdiom[1].lastAskedAt, 3000);
});

test('loadProgress 损坏数据回退默认值', () => {
  const mem = new Map([['idiom.progress', '{{{坏数据']]);
  const storage = { available: true, backend: { getItem: k => mem.get(k) ?? null, setItem: (k, v) => mem.set(k, v), removeItem: k => mem.delete(k) } };
  assert.deepStrictEqual(App.loadProgress(storage), App.newProgress());
  mem.set('idiom.progress', JSON.stringify({ byIdiom: { 3: { asked: 2, correct: 1, correctStreak: 0, wrongLast: true, lastAskedAt: 9 }, bad: 'x' } }));
  const p = App.loadProgress(storage);
  assert.strictEqual(p.byIdiom[3].asked, 2);
  assert.strictEqual(p.byIdiom.bad, undefined, '非法条目被丢弃');
  assert.strictEqual(App.saveProgress(p, storage), true);
});

test('validateData 检出结构问题', () => {
  assert.strictEqual(App.validateData(null).ok, false);
  const good = { books: [{ id: 'b1', name: 'x' }], idioms: [IDS[0]] };
  assert.strictEqual(App.validateData(good).ok, true);
  const bad = JSON.parse(JSON.stringify(good));
  bad.idioms[0].sentences.pop();
  bad.idioms[0].sentences[0].text = '没有空格';
  const r = App.validateData(bad);
  assert.strictEqual(r.ok, false);
  assert.ok(r.errors.length >= 2);
});
```

- [ ] **Step 2: 跑测试确认全红**

Run: `cd C:\Users\MX-GMKTEC-X1\MyWork\ChineseIdiom && node --test tools/logic.test.mjs`
Expected: FAIL（`成语学习小助手.html` 不存在，readFileSync 抛 ENOENT，全部测试红）

- [ ] **Step 3: 创建 `成语学习小助手.html`（骨架 + 全部 CSS + 逻辑标记区）**

```html
<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>成语学习小助手</title>
<style>
:root{--bg:#f6f8fa;--card:#fff;--ink:#1f2937;--muted:#6b7280;--primary:#2563eb;--primary-soft:#dbeafe;--ok:#16a34a;--ok-soft:#dcfce7;--bad:#dc2626;--bad-soft:#fee2e2;--line:#e5e7eb;--radius:14px;--shadow:0 1px 3px rgba(16,24,40,.08),0 8px 24px rgba(16,24,40,.06)}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font-family:"Segoe UI","PingFang SC","Microsoft YaHei",system-ui,sans-serif;line-height:1.6}
.wrap{max-width:1080px;margin:0 auto;padding:16px}
.top{display:flex;flex-wrap:wrap;align-items:center;gap:12px;justify-content:space-between;margin-bottom:16px}
.top h1{font-size:22px;margin:0}
.tabs{display:flex;gap:8px}
.tab{border:1px solid var(--line);background:var(--card);padding:8px 16px;border-radius:999px;cursor:pointer;font-size:15px;color:var(--ink)}
.tab.active{background:var(--primary);border-color:var(--primary);color:#fff}
.banner{background:var(--bad-soft);color:var(--bad);padding:12px 16px;border-radius:var(--radius);margin-bottom:16px}
.toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:16px}
.filters{display:flex;gap:6px;flex-wrap:wrap}
.fbtn{border:1px solid var(--line);background:var(--card);border-radius:999px;padding:4px 12px;cursor:pointer;font-size:13px;color:var(--muted)}
.fbtn.active{background:var(--primary-soft);border-color:var(--primary);color:var(--primary)}
#search{border:1px solid var(--line);border-radius:999px;padding:6px 14px;font-size:14px;min-width:160px}
.btn{border:1px solid var(--line);background:var(--card);border-radius:10px;padding:8px 16px;cursor:pointer;font-size:14px;color:var(--ink)}
.btn.primary{background:var(--primary);border-color:var(--primary);color:#fff}
.btn.primary:disabled{opacity:.5;cursor:not-allowed}
.btn.ghost{color:var(--muted)}
.btn.big{font-size:16px;padding:12px 24px}
.muted{color:var(--muted);font-size:13px}
.card-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:16px}
.flip-card{perspective:1000px;height:200px;cursor:pointer}
.flip-inner{position:relative;width:100%;height:100%;transform-style:preserve-3d;transition:transform .5s}
.flip-card.flipped .flip-inner{transform:rotateY(180deg)}
.flip-face{position:absolute;inset:0;backface-visibility:hidden;-webkit-backface-visibility:hidden;display:flex;align-items:center;justify-content:center;padding:18px;border-radius:var(--radius);background:var(--card);box-shadow:var(--shadow);text-align:center;overflow:auto}
.flip-front .idiom{font-size:30px;font-weight:700;letter-spacing:2px}
.flip-front .no{margin-top:8px;color:var(--muted);font-size:12px}
.flip-back{transform:rotateY(180deg);align-items:flex-start;justify-content:flex-start;font-size:14px;color:#374151;text-align:left}
.quiz-head{display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:12px;flex-wrap:wrap}
#sentences{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:10px}
.sentence{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:12px 16px;display:flex;gap:10px;align-items:baseline}
.s-no{color:var(--muted);font-size:13px;min-width:24px}
.slot{display:inline-block;min-width:7em;border-bottom:2px dashed var(--primary);margin:0 2px;padding:0 6px;text-align:center;color:var(--primary);font-weight:600;cursor:pointer;border-radius:6px}
.slot.filled{background:var(--primary-soft);border-bottom-style:solid}
.slot.hover{background:var(--primary-soft);outline:2px solid var(--primary)}
.tray-title{margin:16px 0 8px}
.tray{display:flex;flex-wrap:wrap;gap:10px;padding:14px;background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);min-height:60px}
.chip{border:1px solid var(--primary);background:var(--primary-soft);color:var(--primary);border-radius:10px;padding:8px 14px;font-size:16px;cursor:grab;touch-action:none;user-select:none}
.chip.selected{background:var(--primary);color:#fff}
.chip.used,.chip.dragging{opacity:.25}
.chip.used{pointer-events:none}
.chip.ghost{position:fixed;z-index:99;pointer-events:none;transform:translate(-50%,-50%);box-shadow:var(--shadow)}
.score{font-size:20px;background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:16px;margin-bottom:16px;text-align:center}
#review{list-style:none;margin:0 0 16px;padding:0;display:flex;flex-direction:column;gap:10px}
.review-item{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:12px 16px;border-left:4px solid var(--line)}
.review-item.ok{border-left-color:var(--ok)}
.review-item.bad{border-left-color:var(--bad)}
.ans-ok{color:var(--ok)}
.ans-bad{color:var(--bad);text-decoration:line-through}
.r-meaning{margin-top:6px;font-size:13px;color:var(--muted)}
#panel-result{text-align:center}
#review{text-align:left}
#retry{margin:8px auto 24px}
@media (max-width:600px){.flip-front .idiom{font-size:24px}.chip{font-size:15px;padding:6px 10px}.wrap{padding:12px}.flip-card{height:220px}}
</style>
</head>
<body>
<div class="wrap">
  <header class="top">
    <h1>🎓 成语学习小助手</h1>
    <nav class="tabs">
      <button class="tab active" data-view="review" type="button">📖 成语复习</button>
      <button class="tab" data-view="quiz" type="button">✏️ 成语测试</button>
    </nav>
  </header>
  <div id="error-banner" class="banner" hidden></div>

  <main id="view-review">
    <div class="toolbar">
      <div id="book-filter" class="filters"></div>
      <input id="search" type="search" placeholder="搜索成语…">
      <button id="reset-flip" class="btn ghost" type="button">全部翻回正面</button>
      <span id="count" class="muted"></span>
    </div>
    <div id="cards" class="card-grid"></div>
  </main>

  <main id="view-quiz" hidden>
    <section id="panel-quiz">
      <div class="quiz-head">
        <span id="round-info" class="muted"></span>
        <button id="submit" class="btn primary" type="button" disabled>上交答案</button>
      </div>
      <ol id="sentences"></ol>
      <div class="tray-title muted">把下面的成语拖进例句的空格里（手机：先点成语，再点空格）：</div>
      <div id="tray" class="tray"></div>
    </section>
    <section id="panel-result" hidden>
      <div id="score-banner" class="score"></div>
      <ol id="review"></ol>
      <button id="retry" class="btn primary big" type="button">🔄 再一次尝试</button>
    </section>
  </main>
</div>

<script src="data.js"></script>
<script>
/* LOGIC-START */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.IdiomApp = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const PROGRESS_KEY = 'idiom.progress';
  const SETTINGS_KEY = 'idiom.settings';

  function newProgress() { return { byIdiom: {} }; }

  function entryOf(progress, id) {
    if (!progress.byIdiom[id]) {
      progress.byIdiom[id] = { asked: 0, correct: 0, correctStreak: 0, wrongLast: false, lastAskedAt: 0 };
    }
    return progress.byIdiom[id];
  }

  function makeStorage() {
    try {
      const s = typeof window !== 'undefined' ? window.localStorage : null;
      if (!s) throw new Error('no localStorage');
      s.setItem('__probe__', '1'); s.removeItem('__probe__');
      return { available: true, backend: s };
    } catch (e) {
      const mem = new Map();
      return {
        available: false,
        backend: {
          getItem: k => (mem.has(k) ? mem.get(k) : null),
          setItem: (k, v) => { mem.set(k, String(v)); },
          removeItem: k => { mem.delete(k); },
        },
      };
    }
  }

  function loadProgress(storage) {
    try {
      const raw = storage.backend.getItem(PROGRESS_KEY);
      if (!raw) return newProgress();
      const p = JSON.parse(raw);
      if (!p || typeof p !== 'object' || !p.byIdiom || typeof p.byIdiom !== 'object') return newProgress();
      const clean = newProgress();
      for (const [k, v] of Object.entries(p.byIdiom)) {
        if (v && typeof v === 'object' && Number.isFinite(v.asked)) {
          clean.byIdiom[k] = {
            asked: v.asked | 0, correct: v.correct | 0, correctStreak: v.correctStreak | 0,
            wrongLast: !!v.wrongLast, lastAskedAt: Number(v.lastAskedAt) || 0,
          };
        }
      }
      return clean;
    } catch (e) { return newProgress(); }
  }

  function saveProgress(progress, storage) {
    try { storage.backend.setItem(PROGRESS_KEY, JSON.stringify(progress)); return true; }
    catch (e) { return false; }
  }

  function loadSettings(storage) {
    try { return JSON.parse(storage.backend.getItem(SETTINGS_KEY)) || {}; } catch (e) { return {}; }
  }
  function saveSettings(settings, storage) {
    try { storage.backend.setItem(SETTINGS_KEY, JSON.stringify(settings)); return true; } catch (e) { return false; }
  }

  function shuffle(arr, rng = Math.random) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(rng() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  function mulberry32(seed) {
    let t = seed >>> 0;
    return function () {
      t += 0x6D2B79F5;
      let x = Math.imul(t ^ (t >>> 15), 1 | t);
      x ^= x + Math.imul(x ^ (x >>> 7), 61 | x);
      return ((x ^ (x >>> 14)) >>> 0) / 4294967296;
    };
  }

  function pickIdioms(idioms, progress, count = 10, rng = Math.random) {
    const by = progress.byIdiom || {};
    const stat = it => by[it.id] || { asked: 0, correctStreak: 0, wrongLast: false, lastAskedAt: 0 };
    const wrong = [], fresh = [], rest = [];
    for (const it of idioms) {
      const s = stat(it);
      if (s.wrongLast && s.correctStreak < 2) wrong.push(it);
      else if (!s.asked) fresh.push(it);
      else rest.push(it);
    }
    rest.sort((a, b) => stat(a).lastAskedAt - stat(b).lastAskedAt);
    const pool = [...shuffle(wrong, rng), ...shuffle(fresh, rng), ...rest];
    return pool.slice(0, Math.min(count, idioms.length));
  }

  function buildRound(picked, rng = Math.random) {
    return picked.map((item, slot) => {
      const idx = Math.min(item.sentences.length - 1, Math.floor(rng() * item.sentences.length));
      return { slot, idiomId: item.id, sentenceIndex: idx, text: item.sentences[idx].text, answer: item.idiom, meaning: item.meaning };
    });
  }

  function gradeRound(round, userAnswers) {
    const results = round.map(q => ({
      slot: q.slot, idiomId: q.idiomId, text: q.text, meaning: q.meaning,
      correct: userAnswers[q.slot] === q.answer,
      userAnswer: userAnswers[q.slot] == null ? null : userAnswers[q.slot],
      correctAnswer: q.answer,
    }));
    return { results, score: results.filter(r => r.correct).length, total: results.length };
  }

  function updateProgress(progress, round, results, now = Date.now()) {
    for (const r of results) {
      const e = entryOf(progress, r.idiomId);
      e.asked += 1;
      if (r.correct) {
        e.correct += 1; e.correctStreak += 1;
        if (e.correctStreak >= 2) e.wrongLast = false;
      } else {
        e.correctStreak = 0; e.wrongLast = true;
      }
      e.lastAskedAt = now;
    }
    return progress;
  }

  function validateData(data) {
    const errors = [];
    if (!data || typeof data !== 'object') return { ok: false, errors: ['题库不是有效对象'] };
    if (!Array.isArray(data.books) || !data.books.length) errors.push('books 缺失');
    if (!Array.isArray(data.idioms)) errors.push('idioms 缺失');
    else {
      const seen = new Set();
      data.idioms.forEach((it, i) => {
        const tag = 'idioms[' + i + ']';
        if (!it || typeof it !== 'object') { errors.push(tag + ' 不是对象'); return; }
        if (typeof it.id !== 'number' || seen.has(it.id)) errors.push(tag + ' id 重复或非法');
        seen.add(it.id);
        if (typeof it.idiom !== 'string' || !it.idiom) errors.push(tag + ' 成语缺失');
        if (typeof it.meaning !== 'string' || !it.meaning) errors.push(tag + ' 释义缺失');
        if (!Array.isArray(it.sentences) || it.sentences.length !== 3) errors.push(tag + ' 需恰好 3 条例句');
        else it.sentences.forEach((s, j) => {
          if (!s || typeof s.text !== 'string' || s.text.split('____').length !== 2) errors.push(tag + '.sentences[' + j + '] 需含恰好一个____');
          if (!s || (s.source !== 'original' && s.source !== 'new')) errors.push(tag + '.sentences[' + j + '] source 非法');
        });
      });
    }
    return { ok: errors.length === 0, errors };
  }

  return {
    PROGRESS_KEY, SETTINGS_KEY, newProgress, entryOf, makeStorage,
    loadProgress, saveProgress, loadSettings, saveSettings,
    shuffle, mulberry32, pickIdioms, buildRound, gradeRound, updateProgress, validateData,
  };
});
/* LOGIC-END */
</script>
<script>
/* UI 脚本（Task 6/7 填充） */
</script>
</body>
</html>
```

- [ ] **Step 4: 跑测试确认全绿**

Run: `node --test tools/logic.test.mjs`
Expected: 10 tests PASS。红则按断言修 `LOGIC` 区代码（只许改标记区内，HTML/CSS 与逻辑无关）。

- [ ] **Step 5: 浏览器冒烟**

双击打开 `成语学习小助手.html`（或 webapp-testing 打开 `file:///C:/Users/MX-GMKTEC-X1/MyWork/ChineseIdiom/成语学习小助手.html`）：页面不白屏、标题「🎓 成语学习小助手」、两个标签按钮渲染、控制台无报错（视图内容还是空的，Task 6/7 填充）。

- [ ] **Step 6: 删除测试台，提交（证据写进 commit 信息）**

```bash
rm tools/logic.test.mjs
git add 成语学习小助手.html
git commit -m "feat(ui): 成语学习小助手.html 骨架 + 内联核心逻辑（LOGIC 标记区）

测试台 tools/logic.test.mjs（node:test，10 用例）全绿后删除：
- shuffle/pickIdioms（全新、不足10、错题优先、连对2次降权）
- buildRound/gradeRound/updateProgress（streak、wrongLast、asked、lastAskedAt）
- loadProgress 损坏回退 / validateData 结构检出

Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 6: 成语复习视图（翻转卡片）+ 视图切换

把 Task 5 骨架里的 `/* UI 脚本（Task 6/7 填充） */` 占位注释替换为 UI 脚本（一个 IIFE，Task 7 会在其中追加测试视图代码）。

**Files:**
- Modify: `成语学习小助手.html`（第二个 `<script>` 块）

**Interfaces:**
- Consumes: `window.IdiomApp`（Task 5）、`window.IDIOM_DATA`（data.js）
- Produces: IIFE 内的 `switchView(v)`、`esc(s)`、`state`、`fail(msg)`；Task 7 将在同一 IIFE 内追加 `startRound()` 等函数（`switchView` 用 `typeof startRound === 'function'` 守卫，Task 6 阶段点测试标签为空白属预期）

- [ ] **Step 1: 写入 UI 脚本（替换占位注释）**

```html
<script>
(function () {
  'use strict';
  const App = window.IdiomApp;
  const storage = App.makeStorage();
  const data = window.IDIOM_DATA;
  const $ = s => document.querySelector(s);
  const $$ = s => Array.from(document.querySelectorAll(s));
  function esc(s) {
    return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function fail(msg) {
    const b = $('#error-banner');
    b.hidden = false; b.textContent = '⚠️ ' + msg;
    $('#view-review').hidden = true; $('#view-quiz').hidden = true;
    document.querySelector('.tabs').style.display = 'none';
  }
  if (!data) { fail('题库加载失败：请检查 data.js 是否存在且格式正确'); return; }
  const check = App.validateData(data);
  if (!check.ok) { fail('题库数据校验未通过：' + check.errors.slice(0, 5).join('；')); return; }

  const bookName = id => { const b = data.books.find(x => x.id === id); return b ? b.name : id; };

  // ===== 视图切换 =====
  let quizStarted = false;
  function switchView(v) {
    $$('.tab').forEach(b => b.classList.toggle('active', b.dataset.view === v));
    $('#view-review').hidden = v !== 'review';
    $('#view-quiz').hidden = v !== 'quiz';
    App.saveSettings(Object.assign(App.loadSettings(storage), { view: v }), storage);
    if (v === 'quiz' && !quizStarted && typeof startRound === 'function') {
      quizStarted = true; startRound();
    }
  }
  $$('.tab').forEach(btn => btn.addEventListener('click', () => switchView(btn.dataset.view)));

  // ===== 成语复习视图 =====
  const state = { book: App.loadSettings(storage).book || 'all', q: '' };

  function renderFilters() {
    const box = $('#book-filter'); box.innerHTML = '';
    const mk = (id, label) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'fbtn' + (state.book === id ? ' active' : '');
      b.textContent = label;
      b.addEventListener('click', () => {
        state.book = id;
        App.saveSettings(Object.assign(App.loadSettings(storage), { book: id }), storage);
        renderFilters(); renderCards();
      });
      box.appendChild(b);
    };
    mk('all', '全部');
    data.books.forEach(bk => mk(bk.id, bk.name.replace(/^成语练习/, '')));
  }

  function renderCards() {
    const grid = $('#cards'); grid.innerHTML = '';
    const q = state.q.trim();
    const list = data.idioms.filter(it =>
      (state.book === 'all' || it.book === state.book) && (!q || it.idiom.includes(q)));
    $('#count').textContent = '共 ' + list.length + ' 个';
    for (const it of list) {
      const card = document.createElement('div');
      card.className = 'flip-card';
      card.innerHTML =
        '<div class="flip-inner">' +
          '<div class="flip-face flip-front"><div><div class="idiom">' + esc(it.idiom) + '</div>' +
          '<div class="no">#' + it.id + ' · ' + esc(bookName(it.book)) + '</div></div></div>' +
          '<div class="flip-face flip-back"><p>' + esc(it.meaning) + '</p></div>' +
        '</div>';
      card.addEventListener('click', () => card.classList.toggle('flipped'));
      grid.appendChild(card);
    }
  }

  $('#search').addEventListener('input', e => { state.q = e.target.value; renderCards(); });
  $('#reset-flip').addEventListener('click', () =>
    $$('.flip-card.flipped').forEach(c => c.classList.remove('flipped')));
  renderFilters();
  renderCards();
  switchView(App.loadSettings(storage).view === 'quiz' ? 'quiz' : 'review');
})();
</script>
```

- [ ] **Step 2: 浏览器实测（webapp-testing 或手动，file:// 打开）**

Run: 用 webapp-testing（Playwright）打开 `file:///C:/Users/MX-GMKTEC-X1/MyWork/ChineseIdiom/成语学习小助手.html`，或主人浏览器手动双击。

检查清单（逐项确认）：
1. 页面标题（浏览器 tab）为「成语学习小助手」，顶部大标题「🎓 成语学习小助手」；
2. 复习视图默认显示 **80 张卡片**（「共 80 个」）；
3. 点击任一卡片 → 3D 翻转显示释义；再点 → 翻回；
4. 「全部翻回正面」→ 所有翻转的卡片复位；
5. 册筛选：点「（1-20）」→ 只剩 20 张；点「全部」恢复 80 张；
6. 搜索「爱」→ 只剩「爱屋及乌」等含「爱」的卡片；清空恢复；
7. 刷新页面 → 上次选择的册筛选仍生效（settings 持久化）；
8. 点「✏️ 成语测试」标签 → 视图切换（本任务阶段内容空白属预期），再点复习标签正常回来；
9. **控制台全程无报错**；
10. 临时把 `data.js` 改名再刷新 → 显示红色横幅「题库加载失败…」，不白屏（验完改回来）。

- [ ] **Step 3: Commit**

```bash
git add 成语学习小助手.html
git commit -m "feat(ui): 成语复习视图 —— 80 张翻转卡片 + 册筛选 + 搜索

浏览器实测通过：翻转/筛选/搜索/刷新记忆/题库缺失横幅/控制台无报错

Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 7: 成语测试视图（拖拽填空）

在 Task 6 的同一 IIFE 内、`switchView(...)` 初始化行**之前**追加测试视图代码（函数声明会提升，`switchView` 的 `typeof startRound` 守卫自动生效）。

**Files:**
- Modify: `成语学习小助手.html`（UI 脚本块内追加）

**Interfaces:**
- Consumes: `IdiomApp.pickIdioms/buildRound/shuffle/gradeRound/updateProgress/loadProgress/saveProgress/loadSettings/saveSettings`（Task 5 签名）、`esc/$/$$/storage/data/locked 前的 state`（Task 6）
- Produces: `startRound()`（视图首次切入时自动调用）

- [ ] **Step 1: 追加测试视图代码（插在 `switchView(App.loadSettings...)` 行之前）**

```js
  // ===== 成语测试视图 =====
  let round = [], placements = [], selectedChip = null, locked = false;
  let roundNo = App.loadSettings(storage).roundNo || 0;
  let suppressClick = false;
  let dragState = null;

  function startRound() {
    locked = false; selectedChip = null; dragState = null;
    const progress = App.loadProgress(storage);
    const picked = App.pickIdioms(data.idioms, progress, 10);
    round = App.buildRound(picked);
    placements = round.map(() => null);
    roundNo += 1;
    $('#panel-result').hidden = true;
    $('#panel-quiz').hidden = false;
    renderQuiz();
  }

  function renderQuiz() {
    $('#round-info').textContent = '第 ' + roundNo + ' 轮 · ' + round.length + ' 题';
    const ol = $('#sentences'); ol.innerHTML = '';
    round.forEach(q => {
      const parts = q.text.split('____');
      const li = document.createElement('li');
      li.className = 'sentence';
      li.innerHTML = '<span class="s-no">' + (q.slot + 1) + '.</span><span>' + esc(parts[0]) +
        '<span class="slot" data-slot="' + q.slot + '"></span>' + esc(parts[1]) + '</span>';
      ol.appendChild(li);
    });
    const tray = $('#tray'); tray.innerHTML = '';
    App.shuffle(round.map(q => q.answer)).forEach(text => {
      const chip = document.createElement('button');
      chip.type = 'button'; chip.className = 'chip';
      chip.textContent = text; chip.dataset.idiom = text;
      tray.appendChild(chip);
    });
    renderPlacements();
    bindQuizEvents();
  }

  function renderPlacements() {
    $$('.slot').forEach(el => {
      const s = +el.dataset.slot;
      el.textContent = placements[s] || '';
      el.classList.toggle('filled', !!placements[s]);
    });
    $$('#tray .chip').forEach(ch => {
      const used = placements.includes(ch.dataset.idiom);
      ch.classList.toggle('used', used);
      ch.classList.toggle('selected', !used && selectedChip === ch.dataset.idiom);
    });
    const left = placements.filter(p => !p).length;
    const btn = $('#submit');
    btn.disabled = left > 0 || locked;
    btn.textContent = left > 0 ? '上交答案（还差 ' + left + ' 空）' : '上交答案';
  }

  function place(slot, idiomText) {
    if (locked || !idiomText) return;
    const from = placements.indexOf(idiomText);   // 从别的空格挪过来
    if (from !== -1) placements[from] = null;
    placements[slot] = idiomText;
    selectedChip = null;
    renderPlacements();
  }
  function unplace(slot) { if (!locked) { placements[slot] = null; renderPlacements(); } }
  function toggleSelect(text) {
    if (!locked) { selectedChip = selectedChip === text ? null : text; renderPlacements(); }
  }

  function slotUnder(e) {
    const el = document.elementFromPoint(e.clientX, e.clientY);
    return el && el.closest ? el.closest('.slot') : null;
  }
  function onChipPointerDown(e, chip) {
    if (locked || chip.classList.contains('used')) return;
    dragState = { chip: chip, x0: e.clientX, y0: e.clientY, moved: false, ghost: null };
    chip.setPointerCapture(e.pointerId);
  }
  function onChipPointerMove(e) {
    if (!dragState) return;
    if (!dragState.moved) {
      if (Math.hypot(e.clientX - dragState.x0, e.clientY - dragState.y0) <= 8) return;
      dragState.moved = true;                       // 超过 8px → 进入拖拽
      const g = dragState.chip.cloneNode(true);
      g.classList.remove('selected'); g.classList.add('ghost');
      document.body.appendChild(g);
      dragState.ghost = g;
      dragState.chip.classList.add('dragging');
    }
    dragState.ghost.style.left = e.clientX + 'px';
    dragState.ghost.style.top = e.clientY + 'px';
    const under = slotUnder(e);
    $$('.slot').forEach(s => s.classList.toggle('hover', s === under));
  }
  function onChipPointerUp(e) {
    if (!dragState) return;
    const st = dragState; dragState = null;
    $$('.slot').forEach(s => s.classList.remove('hover'));
    if (st.moved) {
      suppressClick = true;                          // 拖拽结束不触发点选
      setTimeout(() => { suppressClick = false; }, 0);
      if (st.ghost) st.ghost.remove();
      st.chip.classList.remove('dragging');
      const slot = slotUnder(e);
      if (slot) place(+slot.dataset.slot, st.chip.dataset.idiom);
      else renderPlacements();
    }
  }

  function bindQuizEvents() {
    $$('#tray .chip').forEach(chip => {
      chip.addEventListener('pointerdown', e => onChipPointerDown(e, chip));
      chip.addEventListener('pointermove', onChipPointerMove);
      chip.addEventListener('pointerup', onChipPointerUp);
      chip.addEventListener('click', () => { if (!suppressClick) toggleSelect(chip.dataset.idiom); });
    });
    $$('.slot').forEach(slot => {
      slot.addEventListener('click', () => {
        if (locked) return;
        const s = +slot.dataset.slot;
        if (placements[s]) unplace(s);
        else if (selectedChip) place(s, selectedChip);
      });
    });
  }

  $('#submit').addEventListener('click', () => {
    if (locked || placements.some(p => !p)) return;
    locked = true;
    const graded = App.gradeRound(round, placements);
    const progress = App.loadProgress(storage);
    App.updateProgress(progress, round, graded.results);
    const saved = App.saveProgress(progress, storage);
    App.saveSettings(Object.assign(App.loadSettings(storage), { roundNo: roundNo }), storage);
    renderResult(graded, saved);
  });

  function renderResult(graded, saved) {
    $('#panel-quiz').hidden = true;
    $('#panel-result').hidden = false;
    const results = graded.results, score = graded.score, total = graded.total;
    const emoji = score === total ? '🏆' : (score >= total * 0.6 ? '👍' : '💪');
    $('#score-banner').innerHTML = emoji + ' 本轮得分 <b>' + score + ' / ' + total + '</b>' +
      (saved ? '' : '<div class="muted">提示：本机浏览器存储不可用，练习进度不会被保存</div>');
    const ol = $('#review'); ol.innerHTML = '';
    results.forEach(r => {
      const parts = r.text.split('____');
      const fill = r.correct
        ? '<b class="ans-ok">' + esc(r.userAnswer) + '</b>'
        : '<b class="ans-bad">' + esc(r.userAnswer || '（未填）') + '</b> → <b class="ans-ok">' + esc(r.correctAnswer) + '</b>';
      const li = document.createElement('li');
      li.className = 'review-item ' + (r.correct ? 'ok' : 'bad');
      li.innerHTML = '<div>' + (r.correct ? '✅' : '❌') + ' ' + esc(parts[0]) + fill + esc(parts[1]) + '</div>' +
        (r.correct ? '' : '<div class="r-meaning">【' + esc(r.correctAnswer) + '】' + esc(r.meaning) + '</div>');
      ol.appendChild(li);
    });
    window.scrollTo(0, 0);
  }

  $('#retry').addEventListener('click', startRound);
```

- [ ] **Step 2: 浏览器实测清单（webapp-testing 或手动）**

1. 切到「✏️ 成语测试」→ 自动出第 1 轮：10 句例句（每句 1 个虚线空格）+ 候选区 10 个成语块（顺序与题目不对应）；
2. 「上交答案」初始置灰，文案「上交答案（还差 10 空）」；每放 1 个数字递减；
3. **点选流**：点成语块（高亮）→ 点空格 → 成语入格、块变淡（used）；再点已填空格 → 成语退回候选区；
4. **拖拽流**（桌面）：按住成语块移动 >8px → 出现跟手虚影、经过空格高亮 → 松手入格；拖到非空格处松手 → 成语回候选区；拖拽结束不触发「点选高亮」；
5. 挪格：把已填入 A 格的成语直接拖/点到 B 格 → A 格清空；
6. 放满 10 空 → 按钮变亮「上交答案」→ 点击 → 结果页：得分横幅（如 👍 7 / 10）、每题 ✅/❌、错题显示「划掉的我的答案 → 正确答案」+【成语】释义；
7. 结果页无法再改动答案（题目区已隐藏）；
8. 点「🔄 再一次尝试」→ 全新一轮 10 题，轮数 +1；**上一轮做错的成语出现在新一轮里**（错题优先）；
9. 刷新页面 → 再进测试视图轮数接续、做错过的成语仍优先出现（localStorage 持久化）；
10. 手机宽度（375px）：布局不横滚、点选流可用；
11. **控制台全程无报错**。

- [ ] **Step 3: Commit**

```bash
git add 成语学习小助手.html
git commit -m "feat(ui): 成语测试视图 —— 拖拽/点选填空 + 判分 + 再一次尝试

浏览器实测通过：11 项清单（拖拽、点选、挪格、上交门控、判分显示、
错题优先重现、刷新持久化、375px 布局、控制台无报错）

Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 8: README + 整体收尾验收（⛔ 含主人实测关卡）

**Files:**
- Create: `README.md`
- Modify（如需）: 任何前面任务遗留问题

- [ ] **Step 1: 写 `README.md`**

```markdown
# 🎓 成语学习小助手

新加坡南洋女子中学 中一高级华文 · 80 个成语复习工具（离线单文件网页）。

## 怎么用

1. 双击打开 `成语学习小助手.html`（任何电脑、任何现代浏览器，无需联网）；
2. **📖 成语复习**：点击卡片翻转看解释；可按册筛选、按成语搜索；
3. **✏️ 成语测试**：每轮 10 句例句，把候选区的成语拖进空格
   （手机：先点成语、再点空格）；放满后「上交答案」看判分和正确答案；
   「再一次尝试」出全新 10 题——做错过的成语会优先再考，连对 2 次算掌握。

## 文件说明

| 文件 | 作用 |
|---|---|
| `成语学习小助手.html` | 工具本体（样式脚本全部内联） |
| `data.js` | 题库（80 成语 × 3 例句）。**想加成语/改例句只改这个文件** |
| `tools/extract_raw.py` | 把 extracted/*.txt 解析成 tools/raw/bN.json |
| `tools/assemble_data.py` | 把 raw + pairings + new_sentences 组装成 data.js |
| `tools/validate_data.py` | 题库完整性校验（改完 data.js 跑一次） |
| `docs/idiom-pairing-review.md` | 41-80 两册的答案配对审查清单 |

## 怎么修改题库

1. 用记事本/VS Code 打开 `data.js`（UTF-8）；
2. 格式：首行 `window.IDIOM_DATA = ` 与结尾 `;` 之间是标准 JSON；
   每个成语 3 条例句，句中空格用 `____`（恰好 4 个下划线、恰好 1 处）；
3. 改完在终端跑：`PYTHONIOENCODING=utf-8 python tools/validate_data.py`，
   看到 `PASS` 才算合格；
4. 刷新网页即可生效。

## 练习记录说明

错题优先、轮数等记录存在浏览器 localStorage（键 `idiom.progress`、
`idiom.settings`），**绑定本机本浏览器**：换电脑工具照常用，但记录不跟随；
换浏览器或清理浏览器数据会清空记录。

## 测试证据（构建期）

- 数据管线：test_extract / test_pairings / test_sentences 三个临时测试台
  全绿后删除（证据见各任务 commit 信息）；
- 题库：`tools/validate_data.py` → PASS: 80 idioms / 240 sentences；
- 网页逻辑：tools/logic.test.mjs（node:test，10 用例）全绿后删除；
- 界面：浏览器实测清单（复习 10 项 / 测试 11 项）全部通过、控制台无报错。
```

- [ ] **Step 2: 终验（完整走一遍）**

1. `PYTHONIOENCODING=utf-8 python tools/validate_data.py` → PASS；
2. `node --check` 不适用于 HTML，改为浏览器双击打开 `成语学习小助手.html`：
   复习视图 → 翻卡 → 测试视图 → 完整做一轮（故意错 2 题）→ 上交 →
   再一次尝试（确认错题重现）→ 刷新页面（确认记录还在）；
3. 375px 手机宽度过一遍两视图；控制台无报错；
4. `git status` 干净（临时测试台都已删除，无散落文件）。

- [ ] **Step 3: ⛔ 主人实测验收关卡**

请主人亲自双击玩一遍（复习 + 至少两轮测试），重点感受：
新编例句是否自然、难度是否合适、拖拽是否顺手。**按主人反馈修改后重验，直到主人满意**。

- [ ] **Step 4: 最终提交**

```bash
git add README.md
git commit -m "docs: README 使用说明 + 整体收尾验收

- 用法/文件说明/题库修改指南/练习记录说明/测试证据汇总
- 终验通过：validate PASS、两视图全流程、375px、控制台无报错
- 主人实测验收：<日期> 通过 / 反馈修改：<列出>

Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

## 计划自检记录（Self-Review）

- **Spec 覆盖**：§3 架构→Task 4/5；§4 题库规格→Task 1-4；§5.1 复习视图→Task 6；§5.2 测试视图→Task 7；§5.3 localStorage→Task 5 逻辑+Task 6/7 调用；§6 界面→Task 5 CSS；§7 错误处理→Task 5（validateData/makeStorage）+Task 6（fail 横幅/不足 10 兜底）；§8 测试策略→各任务 TDD 步骤+validate_data.py；§9 交付物→Task 1-8 各自产出。无遗漏。
- **占位符扫描**：无 TBD/TODO；Task 3 例句为内容创作任务，已给出格式示例、约束清单与自动校验，创作本身按计划分批执行。
- **类型一致性**：`IdiomApp` 各函数签名在 Task 5 Interfaces 中定义，Task 6/7 调用与之一致（`pickIdioms(idioms, progress, count)`、`buildRound(picked, rng)`、`gradeRound(round, placements)` 返回 `{results, score, total}`、`updateProgress(progress, round, results, now)`）；raw JSON 字段（`no/idiom/qno/answer/meaning/sentence`）在 Task 1/2/4 一致；`pairings.json`、`new_sentences.json` 结构在 Task 2/3 定义、Task 4 消费一致。

