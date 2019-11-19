# import os
import sys
import pandas as pd               # 資料處理套件
import matplotlib.pyplot as plt   # 資料視覺化套件

# SaveDirectory = os.getcwd()  # 印出目前工作目錄
# csvForGraph = pd.read_csv(os.path.join(SaveDirectory,
#                                       "CsvForGraph_PortB_20191115.csv")
csvForGraph = pd.read_csv(sys.argv[1])
# csvForGraph.head(3)    # 顯示前3筆資料
plt.figure(figsize=(7, 5))   # 顯示圖框架大小

plt.style.use("ggplot")     # 使用ggplot主題樣式
plt.xlabel("Voltage Sent")    # 設定x座標標題及粗體
plt.ylabel("Voltage Received")  # 設定y座標標題及粗體
plt.title("Power Supply Test",
          fontsize=15)   # 設定標題、字大小及粗體

plt.scatter(csvForGraph["Voltage Sent"],              # x軸資料
            csvForGraph["Voltage Received"],         # y軸資料
            c="m",                                  # 點顏色
            s=50,                                   # 點大小
            alpha=.5,                               # 透明度
            marker="D")                             # 點樣式

plt.savefig(sys.argv[2] + "/" + sys.argv[3] + "_Graph.jpg")   # 儲存圖檔
plt.close()      # 關閉圖表
