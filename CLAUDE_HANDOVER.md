# HW Gaming Lab — Claude 接手文件

> 最後更新：2026-05-18  
> 適用對象：接手此專案的新 Claude session

---

## 1. 專案架構說明

```
hw-gaming-lab/
├── index.html          # 入口，直接 redirect → form.html
├── form.html           # 需求申請表單（主要使用者入口）
├── gantt.html          # 測試排程甘特圖（實驗室內部管理）
├── result.html         # 測試結果查閱（需密碼）
└── mock_results/       # 開發用假資料 xlsx
```

**部署位置**：GitHub Pages  
**Repository**：`https://github.com/wen711002/hw-gaming-lab`  
**線上網址**：`https://wen711002.github.io/hw-gaming-lab/`  
**本機位置**：`D:\GL HTML\HW_Gaming_Lab\`

---

## 2. 技術棧

| 項目 | 內容 |
|------|------|
| 前端 | 純 HTML + CSS + Vanilla JS（無框架） |
| 部署 | GitHub Pages（靜態） |
| 後端 | Microsoft SharePoint（inventec corp） |
| API 中介 | Power Automate HTTP 觸發 flows |
| Excel 解析 | xlsx.full.min.js（CDN） |
| 字型 | Google Fonts（Rajdhani、Noto Sans TC、Share Tech Mono） |
| 版本控制 | Git / GitHub（branch: master） |

**無 Node.js、無 npm、無 build process** — 所有檔案直接部署。

---

## 3. 系統架構與資料流

```
使用者瀏覽器
    │
    ├─ form.html ──── PA_URL ────────→ PA Flow: 建立需求單 → SP Forms List
    │                 PA_QUERY ───────→ PA Flow: 讀取Pending → SP Forms List
    │                 PA_ACCEPT ──────→ PA Flow: 更新狀態   → SP Forms List
    │                 PA_UPLOAD ──────→ PA Flow: 上傳Excel  → SP Results List
    │
    ├─ gantt.html ─── PA_READ ───────→ PA Flow: 讀全部資料 → SP (Members+Settings+Tasks)
    │                 PA_TASK ────────→ PA Flow: 任務CRUD   → SP Tasks List
    │                 PA_MEMBER ──────→ PA Flow: 成員CRUD   → SP Members List
    │                 PA_SETTINGS ────→ PA Flow: 設定更新   → SP Settings List
    │
    └─ result.html ── PA_QUERY_RESULTS→ PA Flow: 讀結果索引 → SP Results List
                      PA_GET_FILE ────→ PA Flow: 取得檔案   → SP Document Library
                      PA_UPLOAD ──────→ PA Flow: 上傳結果   → SP Results List
```

**所有 SharePoint 操作都必須透過 PA Flow 中介**（避免 CORS 與 SP 認證問題）。

---

## 4. API 設計（PA Flow URLs）

### form.html

| 變數 | Workflow ID | 功能 | 方法 |
|------|------------|------|------|
| `PA_QUERY` | `f0be976210fd46378a403ee41435a05d` | 讀取所有 Pending 需求單 | POST `{}` |
| `PA_URL` | `f72524dfbf1649b69aace89bc5b37b57` | 建立新需求單 | POST payload |
| `PA_ACCEPT` | `33b3ad8cd52c4f84964ba7a7f30bb5f3` | 接受/更新表單狀態 | POST |
| `PA_UPLOAD` | `2dc7a41ee5c045548f80940d567f5a90` | 上傳結果 Excel | POST |

### gantt.html

| 變數 | Workflow ID | 功能 | 方法 |
|------|------------|------|------|
| `PA_READ` | `ec95713a15c74114ac054f69c2e1c9c2` | 讀取全部資料（成員/設定/任務） | GET |
| `PA_TASK` | `2d8ce514aa3a49118700bfb72a0c0bfd` | 任務 CRUD | POST |
| `PA_MEMBER` | `a0788872751444548b99442bef704a18` | 成員 CRUD | POST |
| `PA_SETTINGS` | `8eb310bc23884b108a4017b7d77901a8` | 更新設定 | POST |

### result.html

| 變數 | Workflow ID | 功能 |
|------|------------|------|
| `PA_QUERY_RESULTS` | `36652175c1ee4a4e98ef6708ec5d02f6` | 讀取結果索引清單 |
| `PA_GET_FILE` | `30ae0af6532742de8324b0c1cab152d7` | 取得 Excel 檔案（回傳 base64） |
| `PA_UPLOAD` | `2dc7a41ee5c045548f80940d567f5a90` | 上傳測試結果 |

---

## 5. SharePoint 資料結構（欄位對照）

### Forms List（需求單）
`guid: 38d74655-7bde-4395-8b7b-f609e48795dd`

| 欄位 | 類型 | 說明 |
|------|------|------|
| Title | Text | 等同 ProjectCode |
| FormType | Text | `新建` / `複測` |
| Dept | Text | 提需部門 |
| Requester | Text | 申請人 |
| RequesterEmail | Text | 申請人信箱 |
| ProductLine | Text | Strix / TUF / Consumer |
| SubSeries | Text | Intel / AMD / Vivobook |
| ProjectName | Text | 專案名稱（長描述） |
| ProjectCode | Text | 專案代號（短碼） |
| Phase | Text | ER1/ER2/ER3/PR1/PR2/PR3/MP（自由輸入） |
| FanMode | Text | Performance / Turbo / Silent |
| ScheduleStart | DateTime | 測試開始日 |
| ScheduleEnd | DateTime | 測試結束日 |
| ReportDeadline | DateTime | 報告需求日（至少 ScheduleStart+5天） |
| SKUList | MultiLineText | 格式：`1. SKU-XXXX CPU:i7 GPU:RTX \| 2. ...` |
| BatteryItems | MultiLineText | 選取的電池測項，換行分隔 |
| BenchmarkItems | MultiLineText | 選取的效能測項 |
| GamingFPSItems | MultiLineText | 選取的FPS測項 |
| Notes | MultiLineText | 備註 |
| RetestReason | Text | 複測原因 |
| Status | Text | `Pending` / `Accepted` |
| SubmittedDate | DateTime | 送出時間 |

> ⚠️ PA Read Flow（PA_QUERY）的 `$select` 需包含 `SKUList,ReportDeadline`，否則前端讀不到。

### Tasks List（甘特任務）

| 欄位 | 類型 | 說明 |
|------|------|------|
| PersonId | Number | Members List 的 SP ID |
| ProjectCode | Text | 專案代號 |
| SKU | Text | SKU 型號 |
| TestType | Choice | `Battery Life` / `Benchmark` / `Gaming FPS` / `Mixed` / `Retest` / `Fail/Repair` |
| Machines | Number | 使用機台數 |
| StartDate | DateTime | 任務開始日 |
| FailMode | Boolean | 是否為 Fail/Repair 模式 |
| Note | Text | 備註 |
| Status | Text | `Planned` 等 |
| ActualHours | Number | 實際工時 |
| RetestOf | Number | 複測來源任務 ID |

### Members List（成員）

| 欄位 | 說明 |
|------|------|
| Title | 姓名（同時作為 ID） |
| Initials | 縮寫（1–2字）|
| Color | 顏色 hex |
| Active | Boolean |

### Settings List（設定）

| 欄位 | 說明 |
|------|------|
| HPD | Hours Per Day（每日工作時數，預設5） |
| MaxMachinesProject | 單專案最大機台數 |
| MaxMachinesPerson | 單人最大機台數 |
| TCFGJson | JSON 字串，存放各測試類型參數 |

---

## 6. 已完成功能

### form.html
- [x] 多步驟表單（模式選擇 → 基本資料 → 測試項目 → SKU清單 → 確認送出）
- [x] 新建 / 複測 兩種模式
- [x] 目前接單狀態 header（統計案子數、SKU 機台數、時間軸）
- [x] SKU 列表動態新增/刪除列
- [x] 測試項目卡片（Battery/Benchmark/Gaming FPS）
  - 點擊卡片 = 全選/全取消該類別所有子項
  - 預設全部選取
  - 子面板永遠可見
- [x] 報告需求日驗證（須 ≥ 測試開始日 + 5 天）
- [x] 預覽確認頁
- [x] 送出後寄信（PA flow 處理）
- [x] Admin 後台（密碼：`GL#31#`，由 result.html 跳轉）
- [x] 複測查詢（從 SP 帶入既有資料）
- [x] PA_QUERY 讀取接單狀態（需 SP 回傳 SKUList+ReportDeadline）

### gantt.html
- [x] 甘特圖可視化（60/90/120天 3種檢視）
- [x] 從 SharePoint 讀取所有資料（loadAllData）
- [x] 新增/編輯/刪除任務（PA_TASK CRUD）
- [x] 成員管理（新增/移除，PA_MEMBER）
- [x] 設定管理（HPD、最大機台數、TCFG 參數，PA_SETTINGS）
- [x] Busy overlay（資料同步中遮罩）
- [x] TYPE_MAP（SP Choice label ↔ TCFG key 轉換）
- [x] 中文 UI（但側邊欄保持英文）

### result.html
- [x] 密碼保護
- [x] 單一專案查詢 / 比較模式
- [x] 從 SP 讀取結果索引
- [x] Excel 檔案讀取與渲染（xlsx.js）
- [x] 上傳測試結果 Excel
- [x] 下載原始 Excel

---

## 7. 未完成 TODO

- [ ] **PA Read Flow 補 SKUList/ReportDeadline 欄位**  
  修改「回傳pending資料至Netlify」flow 的 `$select` URI，加上 `,SKUList,ReportDeadline`（目前只有基本欄位，導致接單狀態的機台數只能 fallback 為每筆1台）

- [ ] **gantt.html 實際端對端測試**  
  loadAllData 已寫完，CRUD payload 欄位已對齊 PA flow PascalCase，但尚未完整驗證 SP 資料正確寫入/讀回

- [ ] **form.html Admin 後台功能確認**  
  接受需求單（PA_ACCEPT）、上傳 Excel（PA_UPLOAD）的完整流程

- [ ] **result.html PROJECT_DATA 改為動態讀取**  
  目前 `PROJECT_DATA` 是硬寫的階層資料（ProductLine→SubSeries→Project），長期應從 SP 動態取得

- [ ] **表單 ScheduleEnd 欄位**  
  payload 中 `ScheduleEnd` 目前送出空字串，應自動計算（依測試類型估算工期）

---

## 8. 關鍵商業邏輯

### 測試類型設定（TCFG）
```js
// gantt.html 中定義，可由 Settings 頁覆寫，存入 SP 的 TCFGJson
TCFG = {
  battery: {m:5,  a:120, f:5,  c:'#4c9ef5', lbl:'Battery Life', lbl2:'電池續航'},
  bench:   {m:1,  a:5,   f:4,  c:'#a855f7', lbl:'Benchmark',    lbl2:'效能測試'},
  fps:     {m:8,  a:20,  f:16, c:'#00e5a0', lbl:'Gaming FPS',   lbl2:'遊戲 FPS'},
  mixed:   {m:14, a:145, f:21, c:'#ffb347', lbl:'Mixed',        lbl2:'混合測試'},
  retest:  {m:5,  a:20,  f:10, c:'#2dd4bf', lbl:'Retest',       lbl2:'重測'},
  fail:    {m:5,  a:0,   f:8,  c:'#ff5c7a', lbl:'Fail/Repair',  lbl2:'失敗／維修'}
}
// m = 機台小時數, a = 分析小時數, f = fail mode 小時數, c = 顏色
```

### TYPE_MAP（SP label 與 TCFG key 互轉）
```js
var TYPE_MAP = {
  'Battery Life':'battery', 'Benchmark':'bench', 'Gaming FPS':'fps',
  'Mixed':'mixed', 'Retest':'retest', 'Fail/Repair':'fail',
  // 也支援 key → key（identity）
  'battery':'battery', 'bench':'bench', ...
};
```

### PA payload 欄位名稱規則（大小寫敏感）
- **Tasks CRUD**：新增/更新用 `PascalCase`（`PersonId`, `ProjectCode`, `TestType`, `StartDate`...）
- **Tasks 刪除**：用小寫 `id`（`{action:'delete', id: spId}`）
- **Settings 更新**：用小寫 `id`（`{id: settingsSpId, HPD:..., TCFGJson:...}`）
- **Members 刪除**：用小寫 `id`

### SKUList 格式
```
1. SKU-XXXX CPU:i7-13700H GPU:RTX4070 | 2. SKU-YYYY CPU:i9 GPU:RTX4090 | 
```
台數計算：`(skuList.match(/\d+\./g)||[]).length || 1`

### 報告需求日規則
- 最少 = `ScheduleStart + 5 天`
- 若尚未填 ScheduleStart，fallback = 今天 + 6 天
- 選了 ScheduleStart 後，`report_due.min` 即時更新

---

## 9. 環境變數 / 設定

無 `.env` 檔案。所有設定都直接寫在 HTML 的 `<script>` 中：

| 項目 | 位置 | 值 |
|------|------|-----|
| Admin 密碼 | result.html `openAdmin()` | `GL#31#` |
| Result 頁密碼 | result.html `checkPW()` | 見原始碼 |
| SP Site | PA Flow 設定 | `https://inventeccorp.sharepoint.com/sites/IEC1-HWGamingLab` |
| PA Flow 環境 | URL 前綴 | `default2ae41f0cacca40f19c6349475ff385.12` |

---

## 10. Build / Deploy 流程

```bash
# 1. 修改本機檔案
#    D:\GL HTML\HW_Gaming_Lab\*.html

# 2. 推上 GitHub（自動部署至 GitHub Pages）
cd "D:\GL HTML\HW_Gaming_Lab"
git add <changed-files>
git commit -m "描述修改內容"
git push

# 3. 等待 1~2 分鐘 GitHub Pages 更新

# 4. 測試（強制清除快取）
# 瀏覽器按 Ctrl+Shift+R
```

**無任何 build step**，直接 push HTML 即可。

---

## 11. 已知 Bug 與注意事項

### ⚠️ PA Read Flow 缺欄位（待修）
「回傳pending資料至Netlify」flow 的 `$select` 未包含 `SKUList`、`ReportDeadline`，導致：
- 接單狀態機台數 fallback 為每筆 1 台（不精確）
- 時間軸的截止日無法顯示

**修法**：在 PA flow URI 的 `$select=...SubmittedDate` 後加上 `,SKUList,ReportDeadline`

### ⚠️ PA flow 欄位名稱大小寫
Tasks CRUD 的新增/更新用 PascalCase，刪除用小寫 `id`，Settings 用小寫 `id`。改動時務必對照 PA flow 截圖確認 token 名稱。

### ⚠️ form.html 的 `sch_start` input type 切換
使用 `onfocus="this.type='date'" onblur="if(!this.value)this.type='text'"` 技巧顯示 placeholder，`onchange="updReportDueMin()"` 會在使用者選日期後觸發。

### ⚠️ result.html PROJECT_DATA 硬寫
目前 ProductLine/SubSeries/ProjectName 階層是硬寫陣列，新增專案需手動維護。

### ⚠️ CORS
所有 SharePoint API 呼叫必須透過 PA Flow 代理，不可直接呼叫 SP REST API（會有 CORS 和 auth 問題）。

---

## 12. Coding Style 與開發規範

- **語言**：繁體中文 UI，程式碼和變數名稱用英文
- **JS**：ES5 語法（`var`, `function`，不用 `const`/`let`/arrow function）— 維持一致
- **CSS 變數**：使用 `:root` 中定義的 `--bg0~4`, `--blue`, `--cyan`, `--green` 等色彩系統
- **設計風格**：深色賽博龐克主題，`--bg0: #0b0e17` 底色
- **字型**：
  - 標題/UI：Rajdhani（英文數字）、Noto Sans TC（中文）
  - 代碼/數據：Share Tech Mono 或 Consolas
- **側邊欄標籤**：保持英文（Form / Result / Schedule / Guide）
- **PA payload**：寫入前先對照 PA flow token 確認大小寫
- **所有寫入操作後**：呼叫 `loadAllData()` 重新載入（gantt.html 模式）
- **檔案結構**：CSS 在 `<style>` 內，JS 在底部 `<script>` 內，無外部 `.js`/`.css` 檔案

---

## 13. 目前進度與下一步建議

### 已完成
- 完整的需求表單（form.html）含 SP 整合
- 甘特圖（gantt.html）含 SP CRUD
- 結果查閱（result.html）
- 中文化 UI

### 立即要做（優先級高）
1. **修 PA Read Flow `$select`**：補上 `SKUList,ReportDeadline` → 機台數才能正確統計
2. **gantt.html 端對端測試**：驗證 loadAllData + addTask/saveEdit/deleteTask 完整流程

### 中期建議
3. **form.html Admin 後台完整測試**：確認 PA_ACCEPT 流程（接受需求、狀態更新）
4. **result.html PROJECT_DATA 動態化**：改從 SP 讀取，避免手動維護
5. **ScheduleEnd 自動計算**：依測試類型 TCFG 估算工期，填入 payload

### 長期建議
6. **需求單狀態追蹤**：Pending → Accepted → InProgress → Done 完整流程
7. **通知機制**：狀態變更時通知申請人（PA flow 已有發信能力）
8. **手機版 RWD**：目前 UI 未針對小螢幕優化
