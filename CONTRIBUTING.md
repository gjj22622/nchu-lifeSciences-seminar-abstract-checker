# 維護交接 SOP

> 給下一屆接手維護本工具的學長姐看。

## 你是誰？

你是中興生科所博士班學長姐，前任維護者把這個 repo 交給你。可能的觸發場景：

1. 老師發了新版「摘要繕寫模式」（學年度更新）
2. 學弟妹回報 bug 或誤判
3. 規則檔的 `valid_until` 即將到期

## 三種常見任務

### 任務 1｜老師發新規範時

```bash
# 1. 拉最新 repo
git clone https://github.com/<owner>/nchu-lifeSciences-seminar-abstract-checker.git
cd nchu-lifeSciences-seminar-abstract-checker

# 2. 把新版「摘要繕寫模式」docx 跟 rules.yaml 比對，找出差異

# 3. 編輯 rules.yaml，更新對應欄位
#    最少要更新：
#      - version: 加版號
#      - last_updated: 改為今天
#      - valid_until: 推到下一個 2 年
#      - spec_version: 改為新學年度標示
#      - sources: 加新版 docx 描述

# 4. 同步更新 README.md 的「規範來源」與標題的學年度

# 5. 更新 規範速查.md（如有規則變動）

# 6. commit
git add rules.yaml README.md 規範速查.md
git commit -m "feat(rules): 依 XXX-X 學年度更新版調整規範"
git push

# 7. 學弟妹下次執行 check_abstract.py 會自動拉到最新規則
```

### 任務 2｜接到 bug 回報時

例：學弟妹說「我的引用明明對的，被報必改」。

```bash
# 1. 跟學弟妹要那份 docx
# 2. 本地跑檢查器，重現問題
python check_abstract.py 出錯的摘要.docx

# 3. 在 check_abstract.py 中找對應的 check 函數，調整邏輯
# 4. 重跑確認修好
# 5. commit + push
git commit -m "fix(check): 修正 XXX 誤判 (回報者 學弟姓名)"
```

### 任務 3｜規則檔即將到期

如果 `valid_until` 還有 3 個月就到期，但你還沒收到老師新版規範：

**選擇 A**：跟老師確認規範未變，把 `valid_until` 推 2 年
```yaml
valid_until: "2030-06-30"
```

**選擇 B**：規範已變但你還沒整理完，先延期 3 個月給自己時間
```yaml
valid_until: "2028-09-30"  # 從 2028-06-30 推到 2028-09-30
```

## 交接給下一任時

當你也要畢業／離開時：

1. 找一位**新任博一或博二學弟妹**接手
2. 在 README.md 的「維護者」表格新增一列
3. 在 commit message 寫：`docs: 維護權交接 給 [姓名] (學號 ...)`
4. 把這份 CONTRIBUTING.md 跟他走過一遍

## Repo 編輯權限

- GitHub repo owner 應該設「組織」或讓繼任者成為 collaborator
- 暫時做法：個人 fork → 老 owner 把新 owner 加為 collaborator → 在 README 改 URL → 大家用新 fork

## 緊急聯絡

如果你完全 stuck：
- 開 GitHub issue，標 `[help wanted]`
- 找黃盟元老師（如果他還在生科所）

## 維護哲學

- **規則優先**：90% 的更新只需改 `rules.yaml`，不要動 `check_abstract.py`
- **保持簡單**：學弟妹安裝門檻就是 `pip install -r requirements.txt`，不要加奇怪依賴
- **回退機制**：永遠保留「無網路也能跑」的能力（不要強制要求遠端規則）
- **過期警告**：是最後的防呆，永遠不要拿掉

---

*謝謝你接下這個小工具，幫所有未來學弟妹省下踩雷的時間。*
