# 系統架構書 — 北歐接力精確體系網站

> 給「下一次回來改這個專案的自己」看的。涵蓋資料流、`build.py` 內部演算法、
> 前端行為、部署方式與已知地雷。改東西前建議先看第 7 節「維護 SOP」。

## 1. 專案是什麼

把《北歐接力精確體系》一書的三份 Excel 叫序表（一方塊／一梅花／高花開叫）與
五篇書籍章節 Markdown，轉成一個純靜態、有 RWD 的網站，部署在 GitHub Pages。

**這不是手寫的網站。** 唯一該手動編輯的是 `原始資料/` 底下的來源檔案；
`網站/` 整個資料夾都是 `build.py` 的產出物，會被覆寫，不要手改裡面的 `.html`。

```
使用者編輯 → 原始資料/*.xlsx、*.md
                  │
                  ▼
            python3 build.py
                  │
                  ▼
        網站/*.html（全部重新產生）
                  │
                  ▼
        git commit + push → GitHub Actions → GitHub Pages
```

## 2. 目錄結構

```
北歐接力精準/
├── build.py                     ← 網站產生器（唯一的邏輯所在地）
├── ARCHITECTURE.md               ← 本文件
├── .github/workflows/pages.yml   ← 推送到 main 就自動部署 Pages
├── 原始資料/                     ← Source of truth，手動維護
│   ├── 一方塊開叫.xlsx / 一梅花開叫.xlsx / 高花開叫.xlsx
│   ├── 北歐接力工具速查.html      ← 獨立手寫頁面（非 build.py 產生，只被複製）
│   ├── 3基本準則/第三章.md + 1.png…13.png
│   ├── 4一方塊開叫/一方塊開叫.md + 1.png…13.png
│   ├── 5高花開叫/高花開叫.md + 1.png…14.png
│   ├── 7一梅花開叫/一梅花開叫.md + 1.png…12.png
│   └── 8二梅花開叫/二梅花開叫.md + 1.png…5.png
└── 網站/                         ← build.py 的輸出，會被整批覆寫
    ├── index.html
    ├── 北歐接力工具速查.html      ← copy_quickref() 從原始資料複製過來的
    ├── assets/style.css, site.js, images/<章節>/*.png
    ├── 叫序資料/一方塊開叫.html, 一梅花開叫.html, 高花開叫.html
    └── 原始書籍檔案/第三章.html, 一方塊開叫.html, 高花開叫.html,
                          一梅花開叫.html, 二梅花開叫.html
```

## 3. `build.py` 總覽

`main()`（第 749 行）依序呼叫：

1. `copy_quickref()` — 把 `原始資料/北歐接力工具速查.html` 原封不動複製進 `網站/`
   （首頁有連結指到它；它在 `網站/` 外面就會變成死連結，見第 8 節教訓 1）。
2. `build_index()` — 產生首頁，卡片連結到下面兩類頁面。
3. `build_xlsx_page()` × 3 — 每個 xlsx 一個頁面，內含多個工作表分頁籤。
4. `build_book_pages()` — 每篇 Markdown 一個頁面。

共用版型（第 35–152 行）：`NAV_ITEMS`（導覽選單資料）、`render_nav()`、
`PAGE_TEMPLATE` / `render_page()`（頁面外殼）、`page_head()`（標題區塊）、
`write_file()`。改導覽選單項目、加新頁面分類，都從這裡改。

## 4. xlsx → HTML 的核心演算法（第 154–545 行）

這是全案最複雜的部分，設計哲學是：**不輸出 Excel 原始底色**，而是把每個
工作表解析成語意結構，再套用網站自己的設計元件。

### 4.1 表頭偵測：`is_header_style()` / `analyze_sheet()`

原始 Excel 裡，表頭列固定是「粗體 + 深藍底 `1F3864`」（`HEADER_NAVY`）。
`analyze_sheet(ws)` 由上往下掃描第一欄，找到第一個符合此樣式的列當作表頭列：

- 表頭列之前的列 → `leading`（標題／叫序脈絡／定義句，前導說明區塊）
- 表頭列之後、且**再次**符合「粗體+深藍底」樣式的列 → `("SUB", 文字)`
  （「說明」工作表裡「接力術語」「本章特有規則」這類小節分隔列就是這樣抓出來的）
- 其餘非空列 → `("DATA", [各欄文字])`
- 資料尾端「只有第一欄有內容」的列 → 收進 `footnotes`（頁尾附註，如
  「＊無接力序列」），用 `while` 從尾端往前收，直到遇到正常資料列為止

**這個偵測完全依賴 Excel 原始樣式，不是看文字內容。** 如果以後有人在 Excel
裡調整了表頭列的底色（不再是 `1F3864`），這個機制會整個失效——`header_row`
會抓不到，退回用第 1 列當表頭，後面全部跑偏。改 Excel 樣式前務必注意這點。

### 4.2 角色分類：`render_sheet_panel()`

依表頭欄位判斷這個工作表該用哪種版面：

| 表頭特徵 | 角色 | Renderer |
|---|---|---|
| `("項目","內容")` | 說明／FAQ 頁 | `render_info()` |
| `("分支","序列／前提")` | 多分支合併在一張表 | `render_merged_branch()` |
| 第一欄是 `例號`，欄位含「方位」 | 示例牌張 | `render_hands()` |
| 第一欄是 `例號`，欄位含「叫者」 | 示例叫序 | `render_auction()` |
| 其他 | 一般叫序階梯表 | `render_ladder()` |

新增一個工作表時，只要表頭符合上述任一形狀，會自動套對 renderer；
如果是全新形狀，`render_sheet_panel()` 要加一個新分支。

### 4.3 跳轉連結：`linkify_refs()` / `extract_branch_code()`

**重要發現**：原始 Excel 資料裡本來就寫死了「→ 表4「A分支」」「見表 9」
「見表 3、4」這類文字——這是可以直接拿來用的既有結構，不用另外設計標記語法。

- `num_to_index`：掃過工作表標題（如「4.A分支(1♠)」）取開頭數字對應到第幾個
  分頁（0-based index）。
- `merged_indices`：哪些分頁是「9.F-J分支」這種一張表裝多個分支的
  （用 4.1 節的表頭特徵判斷）。
- `extract_branch_code(raw)`：從一列的第一欄文字取出分支代號。兩種來源都認得：
  - 純代號欄（表頭是「代號」），值就是單一字母 `A`
  - 混合欄（一梅花 sheet8 那種），值是 `A　2♥`（字母＋全形空白＋叫品）
- `linkify_refs()`：正則 `表\s*(\d+(?:[、,]\s*\d+)*)(?:「([^」]*)」)?` 找出
  「表N」「表N、M」「表N「標籤」」，換成 `<a class="jump-link" data-sheet="sheet-N">`；
  如果目標分頁是 merged 分支表，還會附上 `data-anchor="branch-N-代號"`
  （代號取自「當前這一列自己的代號」，見上一點）。

點擊行為在前端（`site.js`）完成：切到目標分頁籤 + 捲動到對應錨點。
錨點 id 由 `render_merged_branch()` 產生（`branch-{sheet_index+1}-{code}`），
兩邊命名規則必須對得上，改其中一邊記得改另一邊。

### 4.4 文字渲染管線：`render_cell()`

每個儲存格文字都走同一條管線，**順序不能顛倒**：

```
raw text
  → linkify_refs()      在純文字上找「表N」插入 <a> 標籤
  → 用 <[^>]+> 切開文字與標籤
  → 純文字片段各自跑 suit_colorize()（跳脫 HTML＋♠♥♦♣ 上色＋換行變 <br>）
  → 跑 bidify()（用正則把「數字+花色/NT」包成 <span class="bid">）
```

之所以要「先切開標籤再處理純文字」，是因為 `suit_colorize()` 是逐字元跑
`html.escape()`，如果直接對已經含 `<a ...>` 的字串整段下手，會把標籤本身也
當文字跳脫掉，變成畫面上出現一段裂開的標籤原始碼。**這是踩過的雷**
（markdown 轉換那邊也遇過一模一樣的問題，見 `colorize_suits_html()`，
兩個地方是分開兩份邏輯，概念相同）。

### 4.5 Excel 資料本身的怪癖

- 有些字串開頭多一個直引號 `'54Pick-up`（Excel「強制純文字」殘留符號），
  `cell_text()` 會自動砍掉開頭的 `'`。
- 部分表格分組欄（如示例叫序的「例號」）**只有每組第一列有值**，後面列是空白
  仍屬同一組；但示例牌張的「例號」是**每列都重複**。`render_auction()` 和
  `render_hands()` 的分組邏輯因此不一樣，別互相套用。

## 5. 前端行為（`網站/assets/`）

### 5.1 `style.css`

設計 token 沿用 `原始資料/北歐接力工具速查.html` 那份手寫頁的配色系統
（`--baize` 綠、`--brass` 金、Noto Serif/Sans TC + JetBrains Mono），
深色模式靠 `prefers-color-scheme` + `[data-theme]` 雙軌切換。

新增的叫序資料元件都在 `.bidtbl` / `.code-badge` / `.branch-head` /
`.deftable` / `.example-card` / `.handtbl` / `.auctiontbl` 這幾組 class，
全部沒有依賴 Excel 顏色，純粹用既有 token 組出來。

RWD：`.wrap` 限制最大寬度，`.table-scroll` 包住所有表格讓寬表在手機上
可以橫向捲動而不撐爆版面，導覽列 760px 以下收成漢堡選單。

### 5.2 `site.js`

三組獨立行為，都在同一個 IIFE 裡：

1. 手機導覽選單開關（`.nav-toggle` 切 `.site-nav.menu-open`）
2. 導覽下拉選單（點擊切換 `.nav-item.open`，點外面關閉）
3. 叫序資料分頁籤（`[data-sheet-tabs]` 找到面板容器，點 `.sheet-tab` 切換
   `.sheet-panel.is-active`）
4. 分支跳轉連結（`.jump-link` 點擊 → 找對應 `.sheet-tab` 模擬點擊 →
   `scrollIntoView` 到 `data-anchor` 或 `data-sheet` 指的元素）

## 6. Markdown → HTML（`原始書籍檔案/`）

`md_to_html()` 用 `python-markdown`（`tables`, `sane_lists` 擴充）轉換，
之後做三件事：

1. 圖片路徑從裸檔名（如 `1.png`）改指到 `assets/images/<章節>/`
2. `<p><figure>…</figure></p>` 拆掉外層 `<p>`（markdown 會自動把單獨一行的
   圖片包進 `<p>`，但 `<figure>` 是區塊元素，巢狀在 `<p>` 裡不合法）
3. `colorize_suits_html()` 幫花色符號上色（原理同 4.4 節，先切標籤再處理文字，
   避免污染 `alt=` 屬性——這是實際爆炸過的 bug，別再犯）

`BOOK_DOCS`（第 602 行）是五篇文章的清單（來源路徑、標題、圖片資料夾、
簡介），新增文章從這裡加一筆。

## 7. 維護 SOP

改資料的正確流程：

```bash
# 1. 改 原始資料/ 底下的 xlsx 或 md
# 2. 重新產生整個網站（會覆寫 網站/ 底下所有 .html）
python3 build.py

# 3. 检查（可選但建議，尤其是改了 build.py 本身時）
#    - HTML 結構是否合法、內部連結/圖片路徑是否都能解析
#    - 用瀏覽器打開 網站/index.html 目測一下

# 4. commit + push（GitHub Actions 會自動重新部署 Pages）
git add -A
git commit -m "..."
git push origin main
```

**只改 `原始資料/`、不改 `build.py`**：直接重跑 `python3 build.py` 就好，
不用碰前端程式碼。

**要加一份新的叫序資料 xlsx 或新書籍章節**：在 `build.py` 的 `main()` 裡
分別加一行 `build_xlsx_page(...)` 或在 `BOOK_DOCS` 加一筆，並且要去
`NAV_ITEMS`（第 35 行）加對應的導覽選單項目。

新書籍章節還有**第三步，而且 `build.py` 不會幫你做**：把該章的 PNG 複製到
`網站/assets/images/<image_dir>/`。`md_to_html()` 只負責把 markdown 裡的裸檔名
改寫成 `assets/images/<章節>/xxx.png` 這個路徑，**它假設圖片已經放在那裡了**；
沒放也不會報錯——build 會安靜地成功，然後線上整章的圖全部破圖。

```bash
mkdir -p 網站/assets/images/<image_dir>
cp 原始資料/<N章節資料夾>/*.png 網站/assets/images/<image_dir>/
```

（`<image_dir>` 就是 `BOOK_DOCS` 那筆的 `image_dir` 值。）

**要改網站配色／字體／版面**：改 `網站/assets/style.css` 就好，
但因為 `網站/` 整包會被 `build.py` 覆寫——**這份 CSS 檔目前是手動維護、
不是 build.py 產生的**，`build.py` 只會重寫 `.html`，不會動 `style.css`/`site.js`，
所以改了之後不需要、也不會被下一次 `python3 build.py` 蓋掉。

## 8. 已知地雷（都是這次開發實際踩過的）

1. **首頁連到 `網站/` 資料夾外面的檔案會在 GitHub Pages 上變死連結。**
   GitHub Actions 只打包部署 `網站/` 這個資料夾（見 `.github/workflows/pages.yml`
   的 `path: "網站"`），任何指到 `../原始資料/...` 的連結在本機看得到、
   線上一定 404。這也是為什麼 `copy_quickref()` 要把速查頁複製進 `網站/`
   而不是直接連到原始資料夾。**以後再加任何外部檔案連結，先確認目標檔案
   有沒有在 `網站/` 裡。**

2. **HTML id／CSS selector 裡不能有沒跳脫的點號。**
   `document.querySelector('#foo.xlsx-panels')` 裡的 `.` 會被解析成
   class selector，不是字面上的點——即使那個 id 真的叫
   `foo.xlsx-panels`，也選不到。教訓：**組 id 一律用不含 `.`／特殊字元的
   識別字串**（這次改用 `title` 而非帶副檔名的 `xlsx_name`）。

3. **對已經含 HTML 標籤的字串做「逐字元跳脫／文字替換」處理，會炸開標籤。**
   花色上色、跳轉連結插入，都必須「先用 `re.split(r"(<[^>]+>)", ...)`
   把標籤與純文字分開，只處理純文字片段」，不能對整段字串暴力 regex/escape。
   4.4 節、6 節都各自吃過這個虧。

4. **Excel 鎖定檔 `~$*.xlsx`**：開著 Excel 編輯來源檔案時會產生暫存鎖定檔，
   已加進 `.gitignore`（`~$*.xlsx`），不要手動 commit 它。

## 9. 部署架構

- GitHub repo：`https://github.com/Bonki1268/bridge-relay-bidding`（public，
  免費帳號需要 public repo 才能開 GitHub Pages）
- Pages 設定：build source 是「GitHub Actions」（用
  `gh api -X POST repos/.../pages -f build_type=workflow` 設定過一次，
  之後不用再動）
- Workflow：`.github/workflows/pages.yml`，push 到 `main` 或手動
  `workflow_dispatch` 都會觸發，用 `actions/upload-pages-artifact` 打包
  `網站/` 目錄、`actions/deploy-pages` 部署
- 線上網址：`https://bonki1268.github.io/bridge-relay-bidding/`
- 推送需要 GitHub token 有 `workflow` scope（第一次推送因為缺這個 scope
  被拒絕過，用 `gh auth refresh -h github.com -s workflow` 補的）

## 10. 之後可能想做的擴充

- 叫序資料頁目前是「一個 xlsx = 一個頁面 + N 個分頁籤」，如果分頁籤數量
  再增加，可以考慮加一個分頁籤搜尋/篩選框
- `render_ladder()` 目前沒有把「巢狀子接力」（說明頁提到的 A分支 2♦、
  E分支 3♣/3♦ 底下還有一層獨立接力）做視覺上的縮排區分，
  純粹用來源資料原本的列順序平鋪呈現——如果之後想要更精確還原巢狀結構，
  這是 `render_ladder()` 可以加強的方向
- 目前跳轉連結只處理同一個 xlsx 頁面「內部」的分頁跳轉；如果之後想要
  「一梅花開叫」的表格連到「一方塊開叫」頁面的分頁，需要另外設計
  跨頁面連結機制（目前 `num_to_index` 是每個 workbook 各自獨立算的）
