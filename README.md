# 平安住・專案管理系統

- **總網頁（內部）**：https://yuengshih-work.github.io/projects/ ，需管理密碼
- **各案場頁（給業主看）**：https://yuengshih-work.github.io/projects/site/?id=案場代號
  - 短網址也可以：https://yuengshih-work.github.io/projects/案場代號
- **龍德路**：維持原本獨立頁 https://yuengshih-work.github.io/longde/ ，總網頁讀它的後端顯示進度與收款

## 檔案

| 檔案 | 用途 |
|---|---|
| `index.html` | 總網頁：案場清單、可請款提醒、收款總計、新增／設定案場 |
| `site/index.html` | 案場頁範本：總覽、合約文件、平面圖、施工前、施工進度 |
| `assets/common.js` | 共用設定（**API 網址填這裡**）與金額計算 |
| `assets/app.css` | 共用樣式 |
| `404.html` | 短網址轉址 |
| `gas/Code.gs` | Apps Script 後端（不需上傳 GitHub 也可以，放著備份） |

## 部署（帳號 yuengshih@gmail.com）

1. 開 https://script.google.com → 新專案，命名「平安住專案管理系統」，把 `gas/Code.gs` 全部貼上、存檔
2. 上方函式選 `setup` → 執行 → 授權。會自動建立試算表「平安住・專案管理系統」、雲端資料夾，並把龍德路登錄進總網頁
3. 左側齒輪「專案設定」→ 指令碼屬性 → 新增 `ADMIN_PASS`（管理密碼，可和龍德相同）→ **儲存指令碼屬性**
4. 部署 → 新增部署作業 → 類型「網頁應用程式」：執行身分「我」、存取權「所有人」→ 部署，複製 `/exec` 結尾網址
5. 把網址貼到 `assets/common.js` 最上面的 `API`
6. GitHub repo `projects` → Settings → Pages → Branch 選 `main`、資料夾 `/ (root)` → Save

之後修改 Code.gs 要「管理部署作業 → 編輯 → 版本選新版本 → 部署」，網址不變。

## 開新案場

總網頁「＋ 新增案場」→ 填名稱、代號（英文小寫，會成為網址）、未稅總價 → 施工階段與驗收點、付款期程可以從現有案場複製再修改 → 儲存 → 「複製連結」傳給業主。

## 資料

- 收款：每期分「未稅金額」「營業稅」兩筆勾選，各自記錄收款日期
- 可請款：該期勾選的條件（主合約簽署、驗收點合格）都達成且尚未收齊時，總網頁會提示
- 已有照片、紀錄、文件的案場不能刪除，請改狀態為「已結案」
