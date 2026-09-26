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

### 加一整册新成语的流程

1. 把 docx 考卷转成文本放进 `extracted/成语练习（81-100）.txt`；
2. 在 `tools/extract_raw.py` 的 `BOOKS` 列表加一行 `('b5', '成语练习（81-100）', 81, 100, 有答案?)`；
3. 跑 `python tools/extract_raw.py` 生成 `tools/raw/b5.json`；
4. 若该册无答案：补 `tools/pairings.json` 的 b5 配对 + 补 `tools/new_sentences.json`
   的 81-100 新例句（每成语 2 句）；
5. 跑 `python tools/assemble_data.py` + `python tools/validate_data.py`。

## 练习记录说明

错题优先、轮数等记录存在浏览器 localStorage（键 `idiom.progress`、
`idiom.settings`），**绑定本机本浏览器**：换电脑工具照常用，但记录不跟随；
换浏览器或清理浏览器数据会清空记录。

## 测试证据（构建期）

- 数据管线：test_extract / test_pairings / test_sentences 三个临时测试台
  全绿后删除（证据见各任务 commit 信息）；
- 题库：`tools/validate_data.py` → PASS: 80 idioms / 240 sentences；
- 网页逻辑：tools/logic.test.mjs（node:test，10 用例）全绿后删除；
- 界面：浏览器实测清单（复习 10 项 / 测试 11 项）全部通过、控制台无报错；
- 人工：41-80 配对抽查、160 新例句抽读、q16 语病修正，均经主人确认；
- **2026-09-26 主人实测验收通过**（含验收期间的两处布局调整：
  考卷版式「词语库在上」、词语库单行显示）。
