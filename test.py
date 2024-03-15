import tkinter as tk
global custom_value
def get_custom_value():
    
    custom_value = entry.get()
    print("Custom value:", custom_value)
    root.destroy()  # 关闭窗口

# 创建主窗口
root = tk.Tk()
root.title("Custom Value Input")

# 创建文本标签
label = tk.Label(root, text="Enter custom value:")
label.pack()

# 创建文本框
entry = tk.Entry(root)
entry.pack()

# 创建确认按钮
button = tk.Button(root, text="OK", command=get_custom_value)
button.pack()

# 运行主循环
root.mainloop()
