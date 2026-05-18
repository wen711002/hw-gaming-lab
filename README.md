# HW Gaming Lab

> 硬體測試實驗室排程與需求管理系統  
> 最後更新：2026-05-18

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Tech Stack](#2-tech-stack)
3. [Architecture & Data Flow](#3-architecture--data-flow)
4. [Pages & Features](#4-pages--features)
5. [API Reference — Power Automate Flows](#5-api-reference--power-automate-flows)
6. [SharePoint Schema](#6-sharepoint-schema)
7. [Key Business Logic](#7-key-business-logic)
8. [Credentials & Environment](#8-credentials--environment)
9. [Build & Deploy](#9-build--deploy)
10. [Known Issues](#10-known-issues)
11. [Coding Conventions](#11-coding-conventions)
12. [Roadmap](#12-roadmap)

---

## 1. Project Overview

HW Gaming Lab 是 Inventec 硬體測試實驗室的內部管理系統，提供：

- **需求申請**：工程師填寫測試需求單，送出後寄通知信
- **排程管理**：實驗室人員在甘特圖上安排測試任務
- **結果查閱**：上傳 / 瀏覽測試結果 Excel 報告

### Repository

| 項目 | 值 |
|------|----|
| GitHub | `https://github.com/wen711002/hw-gaming-lab` |
| 線上網址 | `https://wen711002.github.io/hw-gaming-lab/` |
| 本機路徑 | `D:\GL HTML\HW_Gaming_Lab\` |
| 預設分支 | `master` |

### File Structure

```
hw-gaming-lab/
├── index.html          # 入口，redirect → form.html
├── form.html           # 需求申請表單（主要使用者入口）
├── gantt.html          # 測試排程甘特圖（實驗室內部）
├── result.html         # 測試結果查閱（密碼保護）
├── README.md           # 本文件
└── mock_results/       # 開發用假資料 .xlsx
```

---

## 2. Tech Stack

| 層級 | 技術 |
|------|------|
| 前端 | 純 HTML5 + CSS3 + Vanilla JS（ES5，無框架） |
| 部署 | GitHub Pages（靜態，無 build process） |
| 後端 / 資料庫 | Microsoft SharePoint（inventec corp） |
| API 中介 | Power Automate HTTP trigger flows |
| Excel 解析 | `xlsx.full.min.js`（CDN） |
| 字型 | Google Fonts（Rajdhani、Noto Sans TC、Share Tech Mono） |

> **重要**：無 Node.js、無 npm、無 bundler。所有 `.html` 檔案直接 push 即部署。

---

## 3. Architecture & Data Flow

```
Browser
  │
  ├── form.html
  │     ├── PA_QUERY   →  [PA] 讀取 Pending 需求單  →  SP: Forms List
  │     ├── PA_URL     →  [PA] 建立新需求單          →  SP: Forms List
  │     ├── PA_ACCEPT  →  [PA] 更新需求單狀態        →  SP: Forms List
  │     └── PA_UPLOAD  →  [PA] 上傳結果 Excel        →  SP: Document Library
  │
  ├── gantt.html
  │     ├── PA_READ    →  [PA] 讀取全部資料          →  SP: Members + Settings + Tasks
  │     ├── PA_TASK    →  [PA] 任務 CRUD             →  SP: Tasks List
  │     ├── PA_MEMBER  →  [PA] 成員 CRUD             →  SP: Members List
  │     └── PA_SETTINGS→  [PA] 設定更新              →  SP: Settings List
  │
  └── result.html
        ├── PA_QUERY_RESULTS →  [PA] 讀取結果索引    →  SP: Results List
        ├── PA_GET_FILE      →  [PA] 取得檔案 base64 →  SP: Document Library
        └── PA_UPLOAD        →  [PA] 上傳測試結果    →  SP: Document Library
```

> 所有 SharePoint 操作**必須透過 PA Flow 代理**，不可直接呼叫 SP REST API（CORS + 認證限制）。

---

## 4. Pages & Features

### 4.1 `form.html` — 需求申請表單

#### 已完成

- 多步驟流程：模式選擇 → 基本資料 → 測試項目 → SKU 列表 → 確認送出
- 新建 / 複測 兩種模式；複測可從 SP 帶入既有資料
- **目前接單狀態** header：統計案子數、SKU 機台數、時間軸
- SKU 列表：動態新增 / 刪除列
- 測試項目卡片（Battery Life / Benchmark / Gaming FPS）
  - 點擊卡片 = 全選 / 全取消該類別子項
  - 子面板永遠可見；預設全部選取
- 報告需求日驗證：須 ≥ 測試開始日 + 5 天，動態更新 `min`
- Admin 後台（密碼 `GL#31#`，由 result.html 跳轉）
- 送出後 PA flow 自動寄通知信

#### 待辦

- `ScheduleEnd` 目前送出空字串，應依測試類型估算自動填入

---

### 4.2 `gantt.html` — 測試排程甘特圖

#### 已完成

- 60 / 90 / 120 天三種檢視切換
- `loadAllData()` 從 SP 讀取所有資料（Members + Settings + Tasks）
- 任務 CRUD（新增 / 編輯 / 刪除，透過 PA_TASK）
- 成員管理（新增 / 移除，透過 PA_MEMBER）
- 設定管理（HPD、最大機台數、TCFG，透過 PA_SETTINGS）
- Busy overlay（資料同步中遮罩）
- TYPE_MAP（SP Choice label ↔ TCFG key 雙向轉換）
- 中文 UI（側邊欄標籤保持英文：Form / Result / Schedule / Guide）

#### 待辦

- 端對端 CRUD 完整驗證（SP 資料寫入後正確讀回）

---

### 4.3 `result.html` — 測試結果查閱

#### 已完成

- 密碼保護入口
- 單一專案查詢 / 多專案比較兩種模式
- 從 SP 動態讀取結果索引（PA_QUERY_RESULTS）
- Excel 檔案讀取與渲染（xlsx.js）
- 上傳 / 下載測試結果 Excel

#### 待辦

- `PROJECT_DATA`（ProductLine → SubSeries → ProjectName 階層）目前硬寫於程式碼，應改為從 SP 動態取得

---

## 5. API Reference — Power Automate Flows

> 所有 PA flow base URL：  
> `https://default2ae41f0cacca40f19c6349475ff385.12.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/`

### 5.1 form.html

| 變數 | Workflow ID | 功能 | HTTP 方法 |
|------|-------------|------|-----------|
| `PA_QUERY` | `f0be976210fd46378a403ee41435a05d` | 讀取所有 Pending 需求單 | POST `{}` |
| `PA_URL` | `f72524dfbf1649b69aace89bc5b37b57` | 建立新需求單 | POST payload |
| `PA_ACCEPT` | `33b3ad8cd52c4f84964ba7a7f30bb5f3` | 接受 / 更新需求單狀態 | POST |
| `PA_UPLOAD` | `2dc7a41ee5c045548f80940d567f5a90` | 上傳結果 Excel | POST |

### 5.2 gantt.html

| 變數 | Workflow ID | 功能 | HTTP 方法 |
|------|-------------|------|-----------|
| `PA_READ` | `ec95713a15c74114ac054f69c2e1c9c2` | 讀取全部甘特資料 | GET |
| `PA_TASK` | `2d8ce514aa3a49118700bfb72a0c0bfd` | 任務 CRUD | POST |
| `PA_MEMBER` | `a0788872751444548b99442bef704a18` | 成員 CRUD | POST |
| `PA_SETTINGS` | `8eb310bc23884b108a4017b7d77901a8` | 設定更新 | POST |

### 5.3 result.html

| 變數 | Workflow ID | 功能 |
|------|-------------|------|
| `PA_QUERY_RESULTS` | `36652175c1ee4a4e98ef6708ec5d02f6` | 讀取結果索引清單 |
| `PA_GET_FILE` | `30ae0af6532742de8324b0c1cab152d7` | 取得 Excel（回傳 base64） |
| `PA_UPLOAD` | `2dc7a41ee5c045548f80940d567f5a90` | 上傳測試結果 |

---

## 6. SharePoint Schema

**SP Site**：`https://inventeccorp.sharepoint.com/sites/IEC1-HWGamingLab`

### 6.1 Forms List（需求單）

GUID：`38d74655-7bde-4395-8b7b-f609e48795dd`

| 欄位 | 類型 | 說明 |
|------|------|------|
| `Title` | Text | 同 ProjectCode |
| `FormType` | Text | `新建` / `複測` |
| `Dept` | Text | 申請部門 |
| `Requester` | Text | 申請人姓名 |
| `RequesterEmail` | Text | 申請人信箱 |
| `ProductLine` | Text | Strix / TUF / Consumer |
| `SubSeries` | Text | Intel / AMD / Vivobook |
| `ProjectName` | Text | 專案名稱（長描述） |
| `ProjectCode` | Text | 專案代號（短碼） |
| `Phase` | Text | ER1/ER2/ER3/PR1/PR2/PR3/MP（自由輸入） |
| `FanMode` | Text | Performance / Turbo / Silent |
| `ScheduleStart` | DateTime | 測試開始日 |
| `ScheduleEnd` | DateTime | 測試結束日（目前送出空值） |
| `ReportDeadline` | DateTime | 報告需求日（≥ ScheduleStart + 5 天） |
| `SKUList` | MultiLineText | `1. SKU-XXXX CPU:i7 \| 2. SKU-YYYY ...` |
| `BatteryItems` | MultiLineText | 選取的電池測項（換行分隔） |
| `BenchmarkItems` | MultiLineText | 選取的效能測項 |
| `GamingFPSItems` | MultiLineText | 選取的 FPS 測項 |
| `Notes` | MultiLineText | 備註 |
| `RetestReason` | Text | 複測原因 |
| `Status` | Text | `Pending` / `Accepted` |
| `SubmittedDate` | DateTime | 送出時間 |

> ⚠️ **PA Read Flow（PA_QUERY）的 `$select` 必須包含 `SKUList,ReportDeadline`**，否則前端讀不到這兩個欄位。目前 flow 尚未加入，需手動更新。

### 6.2 Tasks List（甘特任務）

| 欄位 | 類型 | 說明 |
|------|------|------|
| `PersonId` | Number | Members List 的 SP Item ID |
| `ProjectCode` | Text | 專案代號 |
| `SKU` | Text | SKU 型號 |
| `TestType` | Choice | Battery Life / Benchmark / Gaming FPS / Mixed / Retest / Fail/Repair |
| `Machines` | Number | 使用機台數 |
| `StartDate` | DateTime | 任務開始日 |
| `FailMode` | Boolean | 是否為 Fail/Repair 模式 |
| `Note` | Text | 備註 |
| `Status` | Text | `Planned` 等 |
| `ActualHours` | Number | 實際工時 |
| `RetestOf` | Number | 複測來源任務 ID |

### 6.3 Members List（成員）

| 欄位 | 說明 |
|------|------|
| `Title` | 姓名（同時作為 ID） |
| `Initials` | 縮寫（1–2 字元） |
| `Color` | 顏色 hex（如 `#4c9ef5`） |
| `Active` | Boolean |

### 6.4 Settings List（設定，僅一筆）

| 欄位 | 說明 |
|------|------|
| `HPD` | Hours Per Day，每日工時（預設 5） |
| `MaxMachinesProject` | 單專案最大機台數 |
| `MaxMachinesPerson` | 單人最大機台數 |
| `TCFGJson` | JSON 字串，各測試類型參數（可由 UI 覆寫） |

---

## 7. Key Business Logic

### 7.1 測試類型設定（TCFG）

```js
// 定義於 gantt.html，可由 Settings 頁覆寫後存入 SP TCFGJson
TCFG = {
  battery: { m:5,  a:120, f:5,  c:'#4c9ef5', lbl:'Battery Life', lbl2:'電池續航' },
  bench:   { m:1,  a:5,   f:4,  c:'#a855f7', lbl:'Benchmark',    lbl2:'效能測試' },
  fps:     { m:8,  a:20,  f:16, c:'#00e5a0', lbl:'Gaming FPS',   lbl2:'遊戲 FPS' },
  mixed:   { m:14, a:145, f:21, c:'#ffb347', lbl:'Mixed',        lbl2:'混合測試' },
  retest:  { m:5,  a:20,  f:10, c:'#2dd4bf', lbl:'Retest',       lbl2:'重測' },
  fail:    { m:5,  a:0,   f:8,  c:'#ff5c7a', lbl:'Fail/Repair',  lbl2:'失敗／維修' }
}
// m = 機台小時數  a = 分析小時數  f = fail mode 小時數  c = 顏色
```

### 7.2 TYPE_MAP（SP label ↔ TCFG key）

```js
var TYPE_MAP = {
  'Battery Life':'battery', 'Benchmark':'bench', 'Gaming FPS':'fps',
  'Mixed':'mixed', 'Retest':'retest', 'Fail/Repair':'fail',
  'battery':'battery', 'bench':'bench', 'fps':'fps',   // identity pass-through
  'mixed':'mixed', 'retest':'retest', 'fail':'fail'
};
```

### 7.3 PA Payload 欄位名稱規則（大小寫敏感）

| 操作 | 規則 | 範例 |
|------|------|------|
| Tasks 新增 / 更新 | PascalCase | `PersonId`, `ProjectCode`, `TestType`, `StartDate` |
| Tasks 刪除 | 小寫 `id` | `{ action:'delete', id: spId }` |
| Settings 更新 | 小寫 `id` | `{ id: settingsSpId, HPD:..., TCFGJson:... }` |
| Members 刪除 | 小寫 `id` | `{ action:'delete', id: person.spId }` |

### 7.4 SKUList 格式與機台數計算

```
// 格式（多台以 | 分隔）
"1. SKU-XXXX CPU:i7-13700H GPU:RTX4070 | 2. SKU-YYYY CPU:i9 | "

// 計算台數
var cnt = (skuList.match(/\d+\./g) || []).length || 1;
```

### 7.5 報告需求日邏輯

```
最早可選日期 = ScheduleStart + 5 天
若 ScheduleStart 尚未填寫，fallback = 今天 + 6 天
選了 ScheduleStart 後，report_due.min 即時更新（updReportDueMin()）
若已選的 report_due 因 ScheduleStart 修改而變得不合法，自動清空並提示
```

### 7.6 資料載入模式（gantt.html）

所有寫入操作成功後，一律呼叫 `loadAllData()` 重新從 SP 載入全部資料，確保畫面與 SP 同步。

---

## 8. Credentials & Environment

> ⚠️ 以下資訊屬內部使用，請勿公開。

| 項目 | 值 | 位置 |
|------|-----|------|
| Admin 後台密碼 | `GL#31#` | `result.html` `openAdmin()` |
| Result 頁入口密碼 | 見原始碼 `checkPW()` | `result.html` |
| SP 環境 | `default2ae41f0cacca40f19c6349475ff385.12` | PA Flow URL 前綴 |
| SP Site | `https://inventeccorp.sharepoint.com/sites/IEC1-HWGamingLab` | PA Flow 設定 |

---

## 9. Build & Deploy

此專案**無 build process**，修改 HTML 後直接 push 即部署。

```bash
# 修改本機檔案
# D:\GL HTML\HW_Gaming_Lab\*.html

# 推上 GitHub（自動觸發 GitHub Pages 更新）
git add <changed-files>
git commit -m "簡述修改內容"
git push

# 等待約 1~2 分鐘，GitHub Pages 更新後
# 瀏覽器按 Ctrl+Shift+R 強制清除快取重整
```

---

## 10. Known Issues

### 🔴 PA Read Flow 缺少欄位（待修，優先）

**問題**：「回傳pending資料至Netlify」flow 的 `$select` URI 未包含 `SKUList`、`ReportDeadline`。  
**影響**：接單狀態的機台數 fallback 為每筆 1 台；時間軸無截止日。  
**修法**：在 flow URI 的 `$select=...SubmittedDate` 後補上 `,SKUList,ReportDeadline`。

### 🟡 PA Payload 欄位名稱大小寫陷阱

Tasks CRUD 的新增/更新用 PascalCase，刪除/Settings 更新用小寫 `id`。  
改動 PA flow 時需逐一核對 token 名稱，稍有出入 SP 就會忽略該欄位。

### 🟡 sch_start input type 切換機制

```html
<!-- placeholder 透過 type 切換實現，需搭配 onchange 觸發日期驗證 -->
onfocus="this.type='date'" onblur="if(!this.value)this.type='text'" onchange="updReportDueMin()"
```

### 🟡 result.html PROJECT_DATA 硬寫

ProductLine / SubSeries / ProjectName 階層資料目前寫死於原始碼，新增專案須手動更新。

### 🟢 ScheduleEnd 送出空值

`paSubmit()` payload 中 `ScheduleEnd: ''`，目前不影響流程，但 SP 欄位為空。

---

## 11. Coding Conventions

### JavaScript

- 使用 **ES5** 語法（`var`、`function() {}`），不使用 `const` / `let` / 箭頭函數
- 所有 PA 呼叫成功後以 `.then(function(){ loadAllData(); })` 模式重整資料
- DOM 操作直接用 `document.getElementById` / `querySelectorAll`

### CSS

- 顏色系統以 `:root` CSS 變數定義，不在 inline style 中直接寫色碼
- 主要 token：`--bg0~4`、`--blue`、`--cyan`、`--green`、`--amber`、`--red`、`--purple`
- 設計風格：深色賽博龐克，底色 `--bg0: #0b0e17`

### 字型分工

| 用途 | 字型 |
|------|------|
| 標題 / UI 英文數字 | Rajdhani |
| 內文 / 中文 | Noto Sans TC |
| 代碼 / 數據 | Share Tech Mono、Consolas |

### 命名

- 側邊欄頁籤：保持英文（`Form` / `Result` / `Schedule` / `Guide`）
- PA payload key：嚴格對照各 flow 的 token 名稱

### 檔案結構

- CSS 寫在 `<style>` 內
- JS 寫在底部 `<script>` 內
- **無**外部 `.js` / `.css` 檔案（除 CDN）

---

## 12. Roadmap

### 🔴 立即（高優先）

1. **修 PA Read Flow `$select`** — 補 `SKUList,ReportDeadline`，機台數才能正確統計
2. **gantt.html 端對端驗證** — 確認 loadAllData + addTask / saveEdit / deleteTask 完整流程

### 🟡 短期

3. **form.html Admin 後台完整測試** — PA_ACCEPT 流程（接受需求、狀態更新）
4. **ScheduleEnd 自動計算** — 依 TCFG 估算工期，填入 payload

### 🟢 長期

5. **result.html PROJECT_DATA 動態化** — 改從 SP 讀取，避免手動維護
6. **需求單完整狀態流程** — Pending → Accepted → InProgress → Done
7. **自動通知** — 狀態變更時 PA flow 寄信給申請人
8. **RWD 手機版** — 目前 UI 未針對小螢幕優化
