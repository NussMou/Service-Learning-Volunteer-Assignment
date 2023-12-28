# Service-Learning-Volunteer-Assignment
- Motivation : 由於看到助教在選擇學生時有過多不必要的資訊和重複的操作，所以撰寫此腳本。

---

### 前提與限制：
1. 這個腳本會依據欄位名稱和Sheet的名稱和順序抓取元素，所以名稱必須如下：
    1. Sheet 名稱必須是 `Sheet1`。
    2. 以下欄位名稱必須固定：`姓名`、`Email`、`學號`、`志願一`、`志願二`、`志願三`。
    3. 順序必須是：`志願一`後的下一個欄位必須是「理由和自我推薦」（名稱隨意），以此類推。
2. 第一志願基本上就是最後志願，根據112學年上學期的資料，不到10人的最終志願非第一志願。

---

### 使用方式：
1. 在試算表中點擊 `擴充功能` ，然後點擊 `Apps Script`。
2. 把 `script.gs` 直接複製過去。
3. 選擇函式 `createSheetsBasedOnFirstChoice` 並執行。
4. 把沒有緣分的學生的名字，複製到 Input Sheet 中，Buffer Sheet 即會出現他的完整資料。
5. 把沒有緣分的學生刪掉。

---

### 效果
1. 會分割成諸多 Sheet，每個 Sheet 內僅包含第一志願選擇此服學的學生。
2. 另外還有 Buffer 和 Input，分別用來將沒有緣分的學生放入。
    ![截圖 2023-12-28 下午11.23.04](https://hackmd.io/_uploads/SJN17zjw6.png)
    ![截圖 2023-12-28 下午11.20.42](https://hackmd.io/_uploads/SJwIMfiwa.png)

