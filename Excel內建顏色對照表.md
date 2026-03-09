# Excel 內建背景顏色對照表

程式設定儲存格顏色時，**openpyxl 使用「無 # 的 6 碼十六進位」**，格式為 `RRGGBB`（紅綠藍）。

---

## 一、佈景主題色彩（Theme Colors）

畫面上方「佈景主題色彩」的方塊**沒有固定色碼**，會隨活頁簿的**佈景主題**改變。  
在程式中若要用「主題色」，需用 **Theme 索引**（例如 openpyxl 的 `Theme` 與 `tint`），不建議寫死 hex。

| 常見用途     | 主題索引說明 |
|--------------|--------------|
| 背景 1 / 文字 1 等 | 依主題定義，索引 0～9 |
| 各欄深淺變體     | 同一索引 + 不同 tint 值 |

**結論**：若要**固定顏色、不隨主題變**，請用下面的「標準／調色板」色碼。

---

## 二、標準色彩（Standard Colors）— 畫面下方那一排

介面下方「標準色彩」的 10 格，對應 Excel **ColorIndex 調色板**中的其中 10 色。  
以下為常見對應（依畫面由左到右約略對應）：

| 順序 | 顏色名稱（約略） | ColorIndex | Hex（RRGGBB） | 說明 |
|------|------------------|------------|----------------|------|
| 1    | 深紅             | 9          | 800000         | Dark Red |
| 2    | 鮮紅             | 3          | FF0000         | Red |
| 3    | 黃               | 6          | FFFF00         | Yellow |
| 4    | 淺綠             | 4 或 35    | 00FF00 / 99FF99 | Bright Green / 淺綠 |
| 5    | 亮綠             | 10         | 008000         | Green |
| 6    | 淺藍             | 8 或 41    | 00FFFF / 3366FF | Cyan / 淺藍 |
| 7    | 鮮藍             | 5          | 0000FF         | Blue |
| 8    | 深藍             | 11         | 000080         | Dark Blue |
| 9    | 深紫             | 13         | 800080         | Purple |
| 10   | 紫／靛           | 12 或 14   | 808080 / 等    | 依版本可能不同 |

**在 openpyxl 使用**：  
`PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")`  
（黃色，即上表第 3 格。）

---

## 三、調色板 56 色（ColorIndex 1～56）常用色

Excel 預設調色板共 **56 色**（ColorIndex 1～56）。以下列出常用於「填滿」的色碼，方便在程式中使用。

| ColorIndex | Hex (RRGGBB) | 常見名稱 / 用途 |
|------------|----------------|------------------|
| 1  | 000000 | 黑 |
| 2  | FFFFFF | 白 |
| 3  | FF0000 | 紅 |
| 4  | 00FF00 | 鮮綠 |
| 5  | 0000FF | 藍 |
| 6  | FFFF00 | 黃 |
| 7  | FF00FF | 洋紅 |
| 8  | 00FFFF | 青 |
| 9  | 800000 | 深紅 |
| 10 | 008000 | 綠 |
| 11 | 000080 | 深藍 |
| 12 | 808080 | 灰 |
| 15 | 99CCFF | 淺藍（常用） |
| 34 | C0C0C0 | 銀 |
| 35 | 99FF99 | 淺綠 |
| 36 | CCFFFF | 淡青（本專案 BLOCK_FILL_1） |
| 40 | FFCC99 | 淺橘／桃（本專案 MPO_FILL、BLOCK_FILL_2） |
| 41 | 3366FF | 藍 |
| 43 | 99CC00 | 黃綠 |
| 44 | FFCC00 | 金黃 |
| 45 | 969696 | 灰 |
| 46 | 993366 | 紫褐 |
| 48 | 99CCFF | 淺藍 |
| 49 | CC99FF | 淺紫 |
| 50 | FF99CC | 粉紅 |
| 51 | CC99FF | 淡紫 |
| 52 | FFCC99 | 淺橘（同 40） |
| 53 | 99CCFF | 淺藍 |

**注意**：Excel 內部存的是 **BGR**，若你從 VBA 的 `.Interior.Color` 轉成 hex，要自己轉成 RRGGBB；上表已是 **RRGGBB**，可直接給 openpyxl 用。

---

## 四、本專案目前使用的色碼

| 常數         | Hex     | 說明           |
|--------------|---------|----------------|
| HEADER_FILL  | E6D99C  | 標題列 Row 1：金色，輔色4，較淺60% |
| MPO_FILL     | F9CB9C  | A 欄有 MPO# 的儲存格：橙色，輔色2，較淺60% |
| BLOCK_FILL_1 | DAE3F3  | MPO 區塊一（B~K）：藍色，輔色1，較淺80% |
| BLOCK_FILL_2 | FCE4D6  | MPO 區塊二（B~K）：橙色，輔色2，較淺80% |

---

## 五、在 openpyxl 中指定顏色的方式

```python
from openpyxl.styles import PatternFill

# 用 hex（無 #）
fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")

# 用 openpyxl 內建索引（0–64）
from openpyxl.styles.colors import COLOR_INDEX
# COLOR_INDEX 為 tuple，索引 40 即調色板第 40 色
```

建議：**直接使用上表的 Hex (RRGGBB)** 最穩定，不隨 Excel 版本或主題改變。
