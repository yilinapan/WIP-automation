## 一、用途說明

本文件說明如何在 **macOS** 上使用本專案，提供兩種方式：

- **方式 A：.command 腳本** – 類似 Windows 的 `.bat`，雙擊即可「安裝 + 啟動」。
- **方式 B：Automator .app** – 把腳本包成一顆 Mac 應用程式，雙擊 `.app` 即可啟動。

> 兩種方式背後邏輯相同：  
> 第一次執行時自動建立虛擬環境並安裝 `requirements.txt`，之後直接啟動 Streamlit。

---

## 二、前置條件（所有 Mac 共同需要）

1. **macOS 上已安裝 Python 3**
   - 終端機輸入：

   ```bash
   python3 --version
   ```

   - 若沒有版本號，請安裝 Python 3（可用 Homebrew 或到官方網站下載）。

2. **將專案資料夾放在 Mac 上**
   - 例如放在桌面：`~/Desktop/WIP轉貼紙_Gemini`
   - 資料夾內需包含：
     - `app.py`
     - `processor.py`
     - `requirements.txt`
     - 其他說明檔（本檔案、顏色對照表等）

---

## 三、方式 A：使用 `.command` 腳本（一鍵安裝 + 啟動）

### A-1. 建立腳本檔案

1. 在 Mac 上打開「終端機」。
2. 移動到專案資料夾（依實際路徑調整）：

   ```bash
   cd ~/Desktop/WIP轉貼紙_Gemini
   ```

3. 建立檔案 `run_mac.command`，內容如下（可用任何文字編輯器建立）：

   ```bash
   #!/bin/bash
   # 切到專案目錄（如路徑不同請自行修改）
   cd "$HOME/Desktop/WIP轉貼紙_Gemini" || exit 1

   # 如果沒有虛擬環境就自動建立並安裝套件
   if [ ! -d ".venv" ]; then
     echo "第一次使用，正在建立虛擬環境並安裝套件..."
     python3 -m venv .venv || exit 1
     source .venv/bin/activate
     pip install --upgrade pip
     pip install -r requirements.txt
   else
     source .venv/bin/activate
   fi

   # 啟動 Streamlit
   streamlit run app.py
   ```

> 如果專案不在桌面，請把 `cd "$HOME/Desktop/WIP轉貼紙_Gemini"` 那一行改成實際路徑。

### A-2. 讓腳本可以雙擊執行（只需設定一次）

在終端機中，執行：

```bash
cd ~/Desktop/WIP轉貼紙_Gemini
chmod +x run_mac.command
```

### A-3. 之後的使用方式

- **第一次雙擊 `run_mac.command`：**
  - 自動建立 `.venv` 虛擬環境。
  - 自動安裝 `requirements.txt` 中的套件。
  - 啟動 Streamlit 並打開瀏覽器頁面。

- **之後每次雙擊：**
  - 直接啟動虛擬環境與 Streamlit，不再重複安裝套件。

---

## 四、方式 B：使用 Automator 製作 `.app`（選用）

此方式會建立一顆 `WIP轉貼紙.app`，同事只需雙擊該 App 即可啟動，背後仍然執行與 `run_mac.command` 類似的腳本。

### B-1. 準備啟動腳本內容

腳本內容可與 `run_mac.command` 相同，例如：

```bash
#!/bin/bash
cd "$HOME/Desktop/WIP轉貼紙_Gemini" || exit 1

if [ ! -d ".venv" ]; then
  echo "第一次使用，正在建立虛擬環境並安裝套件..."
  python3 -m venv .venv || exit 1
  source .venv/bin/activate
  pip install --upgrade pip
  pip install -r requirements.txt
else
  source .venv/bin/activate
fi

streamlit run app.py
```

同樣，如果專案路徑不同，請調整 `cd` 那一行。

### B-2. 在 Automator 建立 App（需在 Mac 上操作一次）

1. 在 Mac 上打開 **Automator**。
2. 建立新文件，類型選擇：**「應用程式」**。
3. 在左側搜尋欄輸入「Shell」，選擇「執行 Shell 腳本」，拖到右側工作流程區。
4. 在「執行 Shell 腳本」的設定中：
   - Shell 選擇：`/bin/bash`
   - 「傳遞輸入」選擇：`無輸入`
5. 將上面 B-1 的腳本內容貼入文字框中。
6. 儲存為例如：`WIP轉貼紙.app`（建議存放在桌面或應用程式資料夾）。

### B-3. 之後的使用方式

- 同事只要**雙擊 `WIP轉貼紙.app`**：
  - 第一次執行時，會自動建立虛擬環境並安裝套件。
  - 之後執行時，直接啟動 Streamlit。
- 若 macOS 出現「來自不明開發者」警告，可在：
  - 系統設定 → 隱私權與安全性 → 允許打開此 App。

---

## 五、建議使用策略

- **目前階段（內部少數同事使用）**：建議先採用 **方式 A（.command 腳本）**，設定簡單、修改彈性高。
- **未來若希望完全圖形化操作**：可依需要再加上 **方式 B（Automator .app）**，同事只看到一顆 App。

