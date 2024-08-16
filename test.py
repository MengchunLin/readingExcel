import matplotlib.pyplot as plt
import numpy as np

# 定義地層數據
layers = [
    {"name": "SM1", "color": "lightgreen", "points": [(0, 0.9), (6, 1)]},
    {"name": "CL1", "color": "lightblue", "points": [(0, 0.8), (6, 0.9)]},
    {"name": "SM2", "color": "yellow", "points": [(0, 0.6), (3, 0.7), (6, 0.8)]},
    {"name": "CL2", "color": "lightgreen", "points": [(0, 0.55), (6, 0.6)]},
    {"name": "SM3", "color": "orange", "points": [(2, 0.2), (4, 0.5)]},
    {"name": "CL3", "color": "tan", "points": [(0, 0), (6, 0.2)]}
]

# 創建圖形
fig, ax = plt.subplots(figsize=(10, 6))

# 繪製每個地層
for layer in layers:
    x, y = zip(*layer["points"])
    ax.fill_between(x, y, y2=0, color=layer["color"], alpha=0.7, label=layer["name"])

# 繪製鑽孔位置
drill_holes = [0, 1, 2, 3, 4, 5, 6]
for x in drill_holes:
    ax.axvline(x=x, color='yellow', linestyle='-', linewidth=2)

# 設置圖形屬性
ax.set_xlim(0, 6)
ax.set_ylim(0, 1)
ax.set_title("Kringing 繪製土層剖面")
ax.set_xlabel("距離")
ax.set_ylabel("深度")
ax.legend(loc='upper right')

plt.show()