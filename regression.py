# import pandas as pd
# import numpy as np
# import matplotlib.pyplot as plt
# from scipy.optimize import curve_fit
# import xlwings as xw
# import sys

# # 從命令列取得 Excel 路徑
# excel_path = sys.argv[1]
# wb = xw.Book(excel_path)
# sheet = wb.sheets[0]

# # 讀取 A2:B 到最後一列
# last_row = sheet.range("A2").end('down').row
# x = np.array(sheet.range(f"A2:A{last_row}").value)
# y = np.array(sheet.range(f"B2:B{last_row}").value)

# # 擬合函數
# def model_func(x, a, b, c):
#     return a * np.power(np.maximum(x - c, 0), b)

# params, _ = curve_fit(model_func, x, y, p0=[10000, 2, x.min()])

# # 產生擬合曲線資料
# x_fit = np.linspace(x.min(), x.max(), 500)
# y_fit = model_func(x_fit, *params)

# # 畫圖並儲存暫存圖
# fig, ax = plt.subplots(figsize=(8, 5))
# ax.scatter(x, y, color='skyblue', s=20, label='Original Data')
# ax.plot(x_fit, y_fit, color='red', label='Power Fit')
# ax.set_xlabel('Water Level (m)')
# ax.set_ylabel('Storage Volume (m3)')
# ax.set_title('Regression Fit')
# ax.grid(True)
# ax.legend()
# tmp_path = 'fit_plot.png'
# fig.savefig(tmp_path)
# plt.close()

# # 寫入 G5 儲存格的公式文字
# formula = f"y = {params[0]:.3f} * (x - {params[2]:.3f})^{params[1]:.3f}"
# sheet.range("G5").value = formula

# # 插入圖片
# sheet.pictures.add(tmp_path, name="RegressionPlot", update=True, left=sheet.range("G7").left, top=sheet.range("G7").top)

# # 儲存並關閉（可選）
# # wb.save()
# # wb.close()

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import curve_fit
import xlwings as xw
import sys
import os

print("🚀 回歸分析開始...")

try:
    # 從命令列取得 Excel 路徑
    excel_path = sys.argv[1]
    print(f"📂 取得 Excel 檔案路徑：{excel_path}")
    assert os.path.exists(excel_path), "❌ 找不到指定的 Excel 檔案"

    # 開啟 Excel
    wb = xw.Book(excel_path)
    sheet = wb.sheets[0]
    print("📄 已開啟 Excel 並取得第 1 個工作表")

    # 讀取 A2:B 到最後一列
    last_row = sheet.range("A2").end('down').row
    x = np.array(sheet.range(f"A2:A{last_row}").value)
    y = np.array(sheet.range(f"B2:B{last_row}").value)
    print(f"✅ 成功讀取 {len(x)} 筆資料")

    # 擬合函數
    def model_func(x, a, b, c):
        return a * np.power(np.maximum(x - c, 0), b)

    params, _ = curve_fit(model_func, x, y, p0=[10000, 2, x.min()])
    print("📈 完成曲線擬合")

    # 產生擬合曲線資料
    x_fit = np.linspace(x.min(), x.max(), 500)
    y_fit = model_func(x_fit, *params)

    # 畫圖並儲存暫存圖
    fig, ax = plt.subplots(figsize=(8, 5))
    ax.scatter(x, y, color='skyblue', s=20, label='Original Data')
    ax.plot(x_fit, y_fit, color='red', label='Power Fit')
    ax.set_xlabel('Water Level (m)')
    ax.set_ylabel('Storage Volume (m3)')
    ax.set_title('Regression Fit')
    ax.grid(True)
    ax.legend()
    tmp_path = 'fit_plot.png'
    fig.savefig(tmp_path)
    plt.close()
    print(f"🖼️ 圖片已儲存：{tmp_path}")

    # 寫入 G5 儲存格的公式文字
    formula = f"y = {params[0]:.3f} * (x - {params[2]:.3f})^{params[1]:.3f}"
    sheet.range("G5").value = formula

    # 插入圖片
    sheet.pictures.add(tmp_path, name="RegressionPlot", update=True, 
                       left=sheet.range("G7").left, top=sheet.range("G7").top)
    print("📌 已將擬合公式與圖片寫入 Excel")

    print("✅ 回歸分析完成")
    
except Exception as e:
    print("❌ 發生錯誤：", e)

input("🔚 請按任意鍵關閉視窗...")
