# nchu-lifeSciences-seminar-abstract-checker

> 國立中興大學生命科學系研究所 — 博士班專題討論摘要自動檢查器
> 依 **114-1 學年度更新版**「摘要繕寫模式」+ **Zoological Studies** 引用格式

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

---

## 這是什麼？

把寫好的專討摘要 `.docx` 丟進來，自動檢查格式、字數、引用是否符合系上規範，避免被老師抓格式錯誤。

**為什麼存在**：原作者鐘基啓（博二，2026 年）兩次專討摘要都被老師抓到 Zool Stud 引用格式錯誤（期刊縮寫加句點、卷號未粗體、標題 title case、用 et al. 但作者只有 3 位⋯⋯），整理出一套完整檢查清單後，做成自動化工具留給學弟妹。

---

## 三步驟使用

### 步驟 1｜首次安裝（只需一次）

```bash
pip install -r requirements.txt
```

需要：Python 3.10+、`python-docx`、`pyyaml`

### 步驟 2｜檢查

**Windows 拖放（推薦）**：把 `.docx` 拖到 `拖放這裡檢查.bat` 上

**命令列**：
```bash
python check_abstract.py 你的摘要.docx
```

### 步驟 3｜照報告修

報告分四等級：
- `[必改]` — 一定要改，老師會抓
- `[提醒]` — 老師可能會抓，看學風
- `[資訊]` — 統計資訊
- `[通過]` — 已合規

---

## 檢查項目

### 版面
- 邊界 2.0 cm × 4 邊
- 標題 16 pt 粗體置中（中、英文）
- 摘要 12 pt、1.5 行距、首行縮入 1 cm、左右對齊
- 字數 300–500（中文計字、英文計詞）

### 內容
- 中、英文標題對應、無冒號（避 AI 味）
- 講者姓名 / 學號 / 報告日期齊全且靠右
- 摘要正文無粗體、無斜體（學名除外）
- AI 味禁用詞偵測（「系統性文獻回顧」「賦能」「comprehensive review」⋯⋯）
- 新名詞定義提醒（可選）

### 引用文獻（Zoological Studies 格式）
- 主要參考整條粗體
- 依 Last name 字母排序
- 期刊縮寫**無句點**（`Remote Sens Environ` ✅，`Remote Sens.` ❌）
- 卷號**粗體**（`Forests **14:**1086`）
- 4+ 作者**列前 3 位 + et al.**
- 不列 DOI（依 114-1 規範）
- 頁碼用 en-dash（–），不用 hyphen（-）
- 標題 sentence case，不要 Title Case

### 一頁限制
- A4 一頁估算（超過提醒雙面印）

---

## 規則更新機制（重要！）

本工具的規則**不是寫死在程式碼裡**，而是放在 `rules.yaml`。執行時：

```
1. 嘗試從 rules_remote_url.txt 指定的 URL 拉最新規則
2. 拉不到 → 用本地快取
3. 還是沒有 → 用內建 rules.yaml（會顯示警告）
```

**規則檔自帶有效期限**（`valid_until: 2028-06-30`）。超過後，檢查器會強制顯示：

> ⚠️ 規則檔已超過有效日期，請向學長姐／系辦／黃老師確認最新規範後再交件

也就是說 — **就算原作者畢業失聯、Gist 下架、學弟妹拿到舊版**，他們仍會被提醒去問人，不會誤交。

### 如何更新規則（給未來維護者）

1. **小改動**：在 GitHub repo 直接改 `rules.yaml`，commit + push
2. **同步更新** `last_updated`、必要時更新 `valid_until`
3. 學弟妹下次執行檢查器時**自動拉到最新規則**，無需重新下載程式

不會 Python 也能維護 — `rules.yaml` 是純文字、人類可讀。

### 重大規範變動（如老師發新版「摘要繕寫模式_updated_2027.docx」）

修改流程：
1. 對照新版逐條更新 `rules.yaml`
2. 更新 `spec_version` 欄
3. 在 commit message 註明：「依 XXX 學年度更新版調整」
4. tag 一個新版本（如 `v2027.09.01`）
5. 更新 `規範速查.md` 與 `README.md` 的學年度標示

---

## 資料夾結構

```
nchu-lifeSciences-seminar-abstract-checker/
├── check_abstract.py        主程式（規則執行引擎，幾乎不變）
├── rules.yaml               所有檢查規則（編輯這個就能更新規範）
├── rules_remote_url.txt     遠端規則 URL（執行時優先抓）
├── requirements.txt         Python 依賴
├── 拖放這裡檢查.bat          Windows 拖放執行
├── README.md                本檔
├── 規範速查.md              人類可讀的完整規範文件
├── LICENSE                  MIT
├── .gitignore               
└── CONTRIBUTING.md          維護交接 SOP（給下一屆學長）
```

---

## 規範來源（三層權威）

| 優先 | 來源 |
|---|---|
| 最高 | **114-1 學年度老師更新版**「摘要繕寫模式」docx |
| 補充 | **Zoological Studies 期刊**：https://zoolstud.sinica.edu.tw/page.php?id=7 |
| 歷史 | 113-2 舊版 PDF（已被取代，僅供對照） |

詳見 `規範速查.md`。

---

## 常見問題

**Q：為什麼要用 GitHub 而不是 zip？**
A：GitHub 讓「規則更新」變成一行 commit，所有學弟妹下次跑檢查就拿到最新版。zip 是死的，三年後沒人知道誰手上的版本是對的。

**Q：完全沒網路也能用嗎？**
A：可以。檢查器會 fallback 到本地 `rules.yaml`，只是規則可能不是最新版。

**Q：未來想做 Web 版？**
A：歡迎接手。本專案授權 MIT，fork 後 `python check_abstract.py` 的核心邏輯可包成 API。但要小心伺服器維運成本。

**Q：可以給其他系所用嗎？**
A：可以。fork 本 repo，編輯 `rules.yaml` 即可。本工具設計就是「規則資料化」，不限於生科所。

**Q：A4 一頁裝不下怎麼辦？**
A：把引用清單行距 1.5 → 1.0；或字級 12 → 11。

**Q：et al. 規則為什麼是「3+ 用 et al.」？**
A：Zool Stud 官網 in-text 規則明說 3+ 用 et al.，且 Vancouver style 慣例如此。但範例只列到 3 位全列。我們採「≤3 全列、4+ 列前 3 + et al.」，如果你老師有不同要求請改 `rules.yaml`。

---

## 致謝

- 黃盟元 副教授（指導教授）— 嚴格的格式要求，逼出這套工具
- 朱彥煒 教授（共同指導）

---

## 維護者

| 屆別 | 維護者 | 期間 | 聯絡 |
|---|---|---|---|
| 1.0 | 鐘基啓（Jacky） | 2026 | gjj22622@gmail.com |
| _ | _下一屆學長姐_ | _ | _ |

接手維護請參閱 `CONTRIBUTING.md`。

---

*v1.0 — 2026-05-03*
