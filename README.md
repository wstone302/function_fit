# 使用 Python 對 Excel 中的水位與蓄水量進行非線性函數擬合

本專案將 Excel 中的水位資料與蓄水量進行非線性回歸，擬合出一條符合下式的函數曲線：

> **y = a × (x - c)^b**

並將：
- 📌 擬合結果的公式寫入 Excel 的 G5 儲存格
- 📈 繪製圖表並插入至 G7 位置

---

## 🛠 系統需求

- Windows 作業系統
- 已安裝 Python 3.7+
- 安裝以下 Python 套件：

```bash
pip install pandas numpy matplotlib scipy xlwings

```bash
專案內容
├── function_fit.xlsm     # 含巨集的 Excel 範例檔案
├── regression.py         # 執行擬合與圖表插入的 Python 腳本
