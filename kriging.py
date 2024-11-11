import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from pathlib import Path
import numpy as np
import matplotlib.ticker as ticker
import importlib
import subprocess
import json
import os
from decimal import Decimal, ROUND_UP

# 運行第一個程式
subprocess.run(["python", "Data_processing.py"])

# 讀取第一個程式生成的處理後檔案路徑
with open("processed_files.xlsx", "r") as f:
    processed_files = [line.strip() for line in f.readlines()]


# 使用處理過的檔案進行後續處理
for file in processed_files:
    df = pd.read_excel(file)


# 自動命名並選擇儲存路徑
def auto_save_file(file_path):
    directory, original_filename = os.path.split(file_path)
    name, ext = os.path.splitext(original_filename)
    new_filename = f"{name}_statistical_depth{ext}"
    save_path = os.path.join(directory, new_filename)

    return save_path

# 統計每種土壤的深度範圍，並計算平均IC（前200筆資料不納入計算）
def calculate_depth_statistics_with_qc_avg(df, original_file_path):
    depth_col = df['Depth (m)']
    type_col = df['合併後']
    ic_col = df['Ic']
    Mark_1 = df['Mark1']
    Mark_2 = df['Mark2']
    Bq = df['Bq']

    # 準備變量來記錄每段土壤的範圍和平均IC值
    result = []
    current_type = type_col.iloc[0]  # 從第201筆資料開始
    start_depth = depth_col.iloc[0]
    ic_values = []

    # 遍歷每一行，從第201筆開始，當遇到土壤類型變化或標記改變時，記錄當前土壤段的範圍
    for i in range(201, len(df)):
        if type_col.iloc[i] != current_type:  # 當類型變化時，記錄當前段的數據
            end_depth = depth_col.iloc[i - 1]
            average_ic = sum(ic_values) / len(ic_values) if ic_values else None
            result.append([current_type, start_depth, end_depth, average_ic])
            current_type = type_col.iloc[i]
            start_depth = depth_col.iloc[i]
            ic_values = []  # 重置IC值列表

        # 僅在條件符合時記錄 Ic 值
        if Mark_2.iloc[i] != '*' and pd.notna(Bq.iloc[i]) and Bq.iloc[i] != 0 and Mark_1.iloc[i] != '*':
            ic_values.append(ic_col.iloc[i])

    # 記錄最後一段土壤的範圍及平均IC值
    end_depth = depth_col.iloc[-1]
    average_ic = sum(ic_values) / len(ic_values) if ic_values else None
    result.append([current_type, start_depth, end_depth, average_ic])

    # 創建 DataFrame 保存結果
    depth_stats_df = pd.DataFrame(result, columns=['Type', 'Upper Depth', 'Lower Depth', 'Average Ic'])

    # 自動保存結果
    save_path = auto_save_file(original_file_path)
    depth_stats_df.to_excel(save_path, index=False)


    return depth_stats_df

for i in range(2):
    calculate_depth_statistics_with_qc_avg(pd.read_excel(processed_files[i]), processed_files[i])

# 如果需要分別讀取第一和第二個檔案
df_1 = calculate_depth_statistics_with_qc_avg(pd.read_excel(processed_files[0]), processed_files[0])
df_2 = calculate_depth_statistics_with_qc_avg(pd.read_excel(processed_files[1]), processed_files[1])
# 儲存一份processed_files.xlsx
data_1 = pd.read_excel(processed_files[0])
data_2 = pd.read_excel(processed_files[1])


# 定義鑽孔位置
borehole_position_1 = 0
borehole_position_2 = 1580.53

weight_1 = 0.5
weight_2 = 0.5

# 定義顏色映射
color_mapping = {
    '1': 'lightsalmon',
    '2': 'lightsteelblue',
    '3': 'plum',
    '4': 'darkkhaki',
    '5': 'burlywood',
}

# 初始化變量
layers = []
legend_labels = set()
predict_borehole = pd.DataFrame()
predict_borehole_data = pd.DataFrame()
# 假設 predict_borehole_data 已初始化為一個空的 DataFrame
predict_borehole_data = pd.DataFrame({
    'Depth (m)': pd.Series(dtype='float64'),
    'qc (MPa)': pd.Series(dtype='float64'),
    'fs (MPa)': pd.Series(dtype='float64'),
    'u (MPa)': pd.Series(dtype='float64'),
    'Soil type': pd.Series(dtype='object')
})

depth_ranges = [(0,60),(60,80),(80,110)]  # 定義深度區間

previous_section_1 = None
previous_section_2 = None
data = []
last_depth = 0
# 對比兩個文件的深度區間尋找相近Ic
for depth_range in depth_ranges:

    start_depth, end_depth = depth_range
    matched_layers_major = set()
    matched_layers_minor = set()
    all_types = set()
    section_df_1 = df_1[(df_1['Upper Depth']>start_depth) & (df_1['Upper Depth'] < end_depth) ].reset_index(drop=True)
    section_df_2 = df_2[(df_2['Upper Depth']>start_depth) & (df_2['Upper Depth'] < end_depth) ].reset_index(drop=True)

    # 確認範圍外下一筆資料是否存在並有相同土壤類型
    next_row_1 = df_1[df_1['Upper Depth'] >= end_depth].iloc[:1]
    next_row_2 = df_2[df_2['Upper Depth'] >= end_depth].iloc[:1]
    last_row_1 = section_df_1.iloc[-1]
    last_row_2 = section_df_2.iloc[-1]

    # 確保我們至少有一筆範圍外的資料可供比較
    if not next_row_1.empty and not next_row_2.empty:
    # 取得範圍外的第一筆資料
        next_soil_type_1 = next_row_1.iloc[0]['Type']
        next_soil_type_2 = next_row_2.iloc[0]['Type']

        # 比較範圍外的土壤類型是否相同
        if next_soil_type_1 == last_row_2['Type']:
            # 如果相同，則將範圍外的資料整行加入範圍內
            section_df_1 = pd.concat([section_df_1, next_row_1.iloc[[0]]], ignore_index=True)
        elif next_soil_type_2 == last_row_1['Type']:
            section_df_2 = pd.concat([section_df_2, next_row_2.iloc[[0]]], ignore_index=True)


    # 刪除與上一區間重複的資料
    if previous_section_1 is not None:
        # 使用merge找出重複的行
        duplicates_1 = pd.merge(section_df_1, previous_section_1, how='inner')
        if not duplicates_1.empty:
            # 刪除重複的行
            section_df_1 = section_df_1[~section_df_1.apply(tuple, 1).isin(duplicates_1.apply(tuple, 1))]
            
    if previous_section_2 is not None:
        # 使用merge找出重複的行
        duplicates_2 = pd.merge(section_df_2, previous_section_2, how='inner')
        if not duplicates_2.empty:
            # 刪除重複的行
            section_df_2 = section_df_2[~section_df_2.apply(tuple, 1).isin(duplicates_2.apply(tuple, 1))]
    
    # 重置索引
    section_df_1 = section_df_1.reset_index(drop=True)
    section_df_2 = section_df_2.reset_index(drop=True)
    print(section_df_1)
    print(section_df_2)
    
    # 保存當前區間的數據作為下一次迭代的previous
    previous_section_1 = section_df_1.copy()
    previous_section_2 = section_df_2.copy()
    
    idx_1 = 0
    idx_2 = 0

    len_1 = len(section_df_1)
    len_2 = len(section_df_2)
    # 選出較短的文件
    major_section = section_df_1 if len_1 < len_2 else section_df_2
    minor_section = section_df_2 if len_1 < len_2 else section_df_1
    major_position = borehole_position_1 if len_1 < len_2 else borehole_position_2
    minor_position = borehole_position_2 if len_1 < len_2 else borehole_position_1
    major_data = data_1 if len_1 < len_2 else data_2
    minor_data = data_2 if len_1 < len_2 else data_1

    soil_type_major= major_section['Type']
    soil_type_minor = minor_section['Type']
    # 如果數據需要清理，可以這樣處理：
    upper_depth_major = major_section['Upper Depth'].astype(float)
    upper_depth_minor = minor_section['Upper Depth'].astype(float)
    lower_depth_major = major_section['Lower Depth'].astype(float)
    lower_depth_minor = minor_section['Lower Depth'].astype(float)
    Ic_major = major_section['Average Ic']
    Ic_minor = minor_section['Average Ic']
    count_major = soil_type_major.value_counts()
    count_minor = soil_type_minor.value_counts()
    # 統計count_1和count_2中所有的soil type
    all_types.update(count_major.index)
    all_types.update(count_minor.index)

    match_layer = 0

    for idx in range(len(major_section)):
        include = []
        interpolation = 100
        flag = False
        for i in range(0, 3):
            if idx + i < len(minor_section) and soil_type_major[idx] == soil_type_minor[idx + i] and idx + i not in matched_layers_minor:
                x = abs(Ic_major[idx] - Ic_minor[idx + i])
                if x < interpolation:
                    interpolation = x
                    match_layer = idx + i
                    flag = True

        # 找深度在match_layer之前未匹配的土層-------
        for i in range(0, match_layer):
            if i < len(minor_section):
                if i not in matched_layers_minor:
                    include.append(i)

        # 匹配match_layer
        if flag:
            layers.append({
                "upper_depth_major": (major_position, upper_depth_major[idx]),
                "lower_depth_major": (major_position, lower_depth_major[idx]),
                "upper_depth_minor": (minor_position, upper_depth_minor[match_layer]),
                "lower_depth_minor": (minor_position, lower_depth_minor[match_layer]),
                "label": soil_type_major[idx],
                "color": color_mapping[str(int(soil_type_major[idx]))],
                "soil_type": soil_type_major[idx],
            })
            # 預測predict_borehole_data的數據
            # 取用layers的數據
            upper_limit = upper_depth_major[idx] * weight_1 + upper_depth_minor[match_layer] * weight_2
            lower_limit = lower_depth_major[idx] * weight_1 + lower_depth_minor[match_layer] * weight_2
            upper_limit = round(upper_limit, 2)
            lower_limit = round(lower_limit, 2)
            print('upper_limit', upper_limit)
            print('lower_limit', lower_limit)
            depth = upper_limit
            if depth - last_depth >= 0.02:
                print('不改變深度', depth)
            else:
                depth = depth + 0.01
            # 初始化變數
            x = 0
            data_major = major_data[(major_data['Depth (m)'] >= upper_limit) & (major_data['Depth (m)'] <= lower_limit)]
            data_minor = minor_data[(minor_data['Depth (m)'] >= upper_limit) & (minor_data['Depth (m)'] <= lower_limit)]
            # 遍歷深度範圍
            while depth < lower_limit:
                # 先檢查索引 x 是否在 data_major 和 data_minor 範圍內
                row_major = data_major.iloc[x] if x < len(data_major) else None
                row_minor = data_minor.iloc[x] if x < len(data_minor) else None

                # 判斷 row_major 和 row_minor 是否有數據，並計算合併或單項數據
                if row_major is not None and row_minor is not None:
                    # 合併數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': (row_major.iloc[1] * weight_1 + row_minor.iloc[1] * weight_2),
                        'fs (MPa)': (row_major.iloc[2] * weight_1 + row_minor.iloc[2] * weight_2),
                        'u (MPa)': (row_major.iloc[3] * weight_1 + row_minor.iloc[3] * weight_2),
                    }
                elif row_major is not None and row_minor is None:
                    # 僅使用 row_major 的數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': row_major.iloc[1],
                        'fs (MPa)': row_major.iloc[2],
                        'u (MPa)': row_major.iloc[3],
                    }
                elif row_minor is not None and row_major is None:
                    # 僅使用 row_minor 的數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': row_minor.iloc[1],
                        'fs (MPa)': row_minor.iloc[2],
                        'u (MPa)': row_minor.iloc[3],
                    }

                # 更新索引和深度
                x += 1
                last_depth = depth
                depth += 0.02
                data.append(row)
                
            matched_layers_major.add(idx)
            matched_layers_minor.add(match_layer)


        
        elif not flag:
            if idx != 0:
                layers.append({
                    "upper_depth_major": (major_position, upper_depth_major[idx]),
                    "lower_depth_major": (major_position, lower_depth_major[idx]),
                    "upper_depth_minor": (minor_position, lower_depth_minor[match_layer]),
                    "lower_depth_minor": (minor_position, lower_depth_minor[match_layer]),
                    "label": soil_type_major[idx],
                    "color": color_mapping[str(int(soil_type_major[idx]))],
                    "soil_type": soil_type_major[idx],
                })
                # 預測predict_borehole_data的數據
                # 取用layers的數據
                upper_limit = upper_depth_major[idx] * weight_1 + lower_depth_minor[match_layer] * weight_2
                lower_limit = lower_depth_major[idx] * weight_1 + lower_depth_minor[match_layer] * weight_2
                upper_limit = round(upper_limit, 2)
                lower_limit = round(lower_limit, 2)
                print('upper_limit', upper_limit)
                print('lower_limit', lower_limit)
                depth = upper_limit
                if depth - last_depth >= 0.02:
                    print('不改變深度', depth)
                else:
                    depth = depth + 0.01
                data_major = major_data[(major_data['Depth (m)'] >= upper_limit) & (major_data['Depth (m)'] <= lower_limit)]
                data_minor = minor_data[(minor_data['Depth (m)'] >= upper_limit) & (minor_data['Depth (m)'] <= lower_limit)]
                # 初始化變數
                x = 0

                # 遍歷深度範圍
                while depth < lower_limit:
                    # 先檢查索引 x 是否在 data_major 和 data_minor 範圍內
                    row_major = data_major.iloc[x] if x < len(data_major) else None
                    row_minor = data_minor.iloc[x] if x < len(data_minor) else None

                    # 判斷 row_major 和 row_minor 是否有數據，並計算合併或單項數據
                    if row_major is not None and row_minor is not None:
                        # 合併數據
                        row = {
                            'Depth (m)': depth,
                            'qc (MPa)': (row_major.iloc[1] * weight_1 + row_minor.iloc[1] * weight_2),
                            'fs (MPa)': (row_major.iloc[2] * weight_1 + row_minor.iloc[2] * weight_2),
                            'u (MPa)': (row_major.iloc[3] * weight_1 + row_minor.iloc[3] * weight_2),
                        }
                    elif row_major is not None and row_minor is None:
                        # 僅使用 row_major 的數據
                        row = {
                            'Depth (m)': depth,
                            'qc (MPa)': row_major.iloc[1],
                            'fs (MPa)': row_major.iloc[2],
                            'u (MPa)': row_major.iloc[3],
                        }
                    elif row_minor is not None and row_major is None:
                        # 僅使用 row_minor 的數據
                        row = {
                            'Depth (m)': depth,
                            'qc (MPa)': row_minor.iloc[1],
                            'fs (MPa)': row_minor.iloc[2],
                            'u (MPa)': row_minor.iloc[3],
                        }

                    # 更新索引和深度
                    x += 1
                    last_depth = depth
                    depth += 0.02
                    data.append(row)
                    
                matched_layers_major.add(idx)

            else:
                # 當主鑽孔第一筆資料的upper_depth_1小於副鑽孔第一筆資料的upper_depth_2
                if upper_depth_minor[idx] < upper_depth_major[0]:
                    layers.append({
                        "upper_depth_major": (major_position, upper_depth_major[0]),
                        "lower_depth_major": (major_position, lower_depth_major[0]),
                        "upper_depth_minor": (minor_position, upper_depth_minor[idx]),
                        "lower_depth_minor": (minor_position, lower_depth_minor[idx]),
                        "label": soil_type_minor[idx],
                        "color": color_mapping[str(int(soil_type_minor[idx]))],
                        "soil_type": soil_type_minor[idx],
                    })
                    # 預測predict_borehole_data的數據
                    # 取用layers的數據
                    upper_limit = upper_depth_major[0] * weight_1 + upper_depth_minor[idx] * weight_2
                    lower_limit = lower_depth_major[0] * weight_1 + lower_depth_minor[idx] * weight_2
                    upper_limit = round(upper_limit, 2)
                    lower_limit = round(lower_limit, 2)
                    print('upper_limit', upper_limit)
                    print('lower_limit', lower_limit)
                    depth = upper_limit
                    if depth - last_depth >= 0.02:
                        print('不改變深度', depth)
                    else:
                        depth = depth + 0.01
                    data_major = major_data[(major_data['Depth (m)'] >= upper_limit) & (major_data['Depth (m)'] <= lower_limit)]
                    data_minor = minor_data[(minor_data['Depth (m)'] >= upper_limit) & (minor_data['Depth (m)'] <= lower_limit)]
                    # 初始化變數
                    x = 0
                    # 遍歷深度範圍
                    while depth < lower_limit:
                        # 先檢查索引 x 是否在 data_major 和 data_minor 範圍內
                        row_major = data_major.iloc[x] if x < len(data_major) else None
                        row_minor = data_minor.iloc[x] if x < len(data_minor) else None

                        # 判斷 row_major 和 row_minor 是否有數據，並計算合併或單項數據
                        if row_major is not None and row_minor is not None:
                            # 合併數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': (row_major.iloc[1] * weight_1 + row_minor.iloc[1] * weight_2),
                                'fs (MPa)': (row_major.iloc[2] * weight_1 + row_minor.iloc[2] * weight_2),
                                'u (MPa)': (row_major.iloc[3] * weight_1 + row_minor.iloc[3] * weight_2),
                            }
                        elif row_major is not None and row_minor is None:
                            # 僅使用 row_major 的數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': row_major.iloc[1],
                                'fs (MPa)': row_major.iloc[2],
                                'u (MPa)': row_major.iloc[3],
                            }
                        elif row_minor is not None and row_major is None:
                            # 僅使用 row_minor 的數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': row_minor.iloc[1],
                                'fs (MPa)': row_minor.iloc[2],
                                'u (MPa)': row_minor.iloc[3],
                            }
                        # 更新索引和深度
                        x += 1
                        last_depth = depth
                        depth += 0.02
                        data.append(row)
                        
                    matched_layers_minor.add(idx)

                    layers.append({
                        "upper_depth_major": (major_position, upper_depth_major[idx]),
                        "lower_depth_major": (major_position, lower_depth_major[idx]),
                        "upper_depth_minor": (minor_position, lower_depth_minor[0]),
                        "lower_depth_minor": (minor_position, lower_depth_minor[0]),
                        "label": soil_type_minor[0],
                        "color": color_mapping[str(int(soil_type_major[0]))],
                        "soil_type": soil_type_minor[0],
                    })
                    # 預測predict_borehole_data的數據
                    # 取用layers的數據
                    upper_limit = upper_depth_major[idx] * weight_1 + lower_depth_minor[0] * weight_2
                    lower_limit = lower_depth_major[idx] * weight_1 + lower_depth_minor[0] * weight_2
                    upper_limit = round(upper_limit, 2)
                    lower_limit = round(lower_limit, 2)
                    print('upper_limit', upper_limit)
                    print('lower_limit', lower_limit)
                    
                    depth = upper_limit
                    if depth - last_depth >= 0.02:
                        print('不改變深度', depth)
                    else:
                        depth = depth + 0.01
                    # 選取data_1和data_2在範圍upper_depth_major、upper_depth_minor、lower_depth_major和lower_depth_minor之間的數據
                    # 使用layers的數據
                    data_major = major_data[(major_data['Depth (m)'] >= upper_limit) & (major_data['Depth (m)'] <= lower_limit)]
                    data_minor = minor_data[(minor_data['Depth (m)'] >= upper_limit) & (minor_data['Depth (m)'] <= lower_limit)]
                    # 初始化變數
                    x = 0

                    # 遍歷深度範圍
                    while depth < lower_limit:
                        # 先檢查索引 x 是否在 data_major 和 data_minor 範圍內
                        row_major = data_major.iloc[x] if x < len(data_major) else None
                        row_minor = data_minor.iloc[x] if x < len(data_minor) else None

                        # 判斷 row_major 和 row_minor 是否有數據，並計算合併或單項數據
                        if row_major is not None and row_minor is not None:
                            # 合併數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': (row_major.iloc[1] * weight_1 + row_minor.iloc[1] * weight_2),
                                'fs (MPa)': (row_major.iloc[2] * weight_1 + row_minor.iloc[2] * weight_2),
                                'u (MPa)': (row_major.iloc[3] * weight_1 + row_minor.iloc[3] * weight_2),
                            }
                        elif row_major is not None and row_minor is None:
                            # 僅使用 row_major 的數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': row_major.iloc[1],
                                'fs (MPa)': row_major.iloc[2],
                                'u (MPa)': row_major.iloc[3],
                            }
                        elif row_minor is not None and row_major is None:
                            # 僅使用 row_minor 的數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': row_minor.iloc[1],
                                'fs (MPa)': row_minor.iloc[2],
                                'u (MPa)': row_minor.iloc[3],
                            }

                        # 更新索引和深度
                        x += 1
                        last_depth = depth
                        depth += 0.02
                        data.append(row)
                        
                    matched_layers_major.add(idx)

                    
                # 當副鑽孔的深度大於主鑽孔的深度
                elif upper_depth_minor[idx] > upper_depth_major[0]:
                    layers.append({
                        "upper_depth_major": (major_position, upper_depth_major[0]),
                        "lower_depth_major": (major_position, lower_depth_major[0]),
                        "upper_depth_minor": (minor_position, upper_depth_minor[idx]),
                        "lower_depth_minor": (minor_position, upper_depth_minor[idx]),
                        "label": soil_type_major[0],
                        "color": color_mapping[str(int(soil_type_major[0]))],
                        "soil_type": soil_type_major[0],
                    })
                                        # 預測predict_borehole_data的數據
                    # 取用layers的數據
                    upper_limit = upper_depth_major[0] * weight_1 + upper_depth_minor[idx] * weight_2
                    lower_limit = lower_depth_major[0] * weight_1 + upper_depth_minor[idx] * weight_2
                    upper_limit = round(upper_limit, 2)
                    lower_limit = round(lower_limit, 2)
                    print(upper_limit, lower_limit)
                    depth = upper_limit
                    if depth - last_depth >= 0.02:
                        print('不改變深度', depth)
                    else:
                        depth = depth + 0.01
                    # 選取data_1和data_2在範圍upper_depth_major、upper_depth_minor、lower_depth_major和lower_depth_minor之間的數據
                    # 使用layers的數據
                    data_major = major_data[(major_data['Depth (m)'] >= upper_limit) & (major_data['Depth (m)'] <= lower_limit)]
                    data_minor = minor_data[(minor_data['Depth (m)'] >= upper_limit) & (minor_data['Depth (m)'] <= lower_limit)]
                    # 初始化變數
                    x = 0

                    # 遍歷深度範圍
                    while depth < lower_limit:
                        # 先檢查索引 x 是否在 data_major 和 data_minor 範圍內
                        row_major = data_major.iloc[x] if x < len(data_major) else None
                        row_minor = data_minor.iloc[x] if x < len(data_minor) else None

                        # 判斷 row_major 和 row_minor 是否有數據，並計算合併或單項數據
                        if row_major is not None and row_minor is not None:
                            # 合併數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': (row_major.iloc[1] * weight_1 + row_minor.iloc[1] * weight_2),
                                'fs (MPa)': (row_major.iloc[2] * weight_1 + row_minor.iloc[2] * weight_2),
                                'u (MPa)': (row_major.iloc[3] * weight_1 + row_minor.iloc[3] * weight_2),
                            }
                        elif row_major is not None and row_minor is None:
                            # 僅使用 row_major 的數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': row_major.iloc[1],
                                'fs (MPa)': row_major.iloc[2],
                                'u (MPa)': row_major.iloc[3],
                            }
                        elif row_minor is not None and row_major is  None:
                            # 僅使用 row_minor 的數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': row_minor.iloc[1],
                                'fs (MPa)': row_minor.iloc[2],
                                'u (MPa)': row_minor.iloc[3],
                            }

                        # 更新索引和深度
                        x += 1
                        last_depth = depth
                        depth += 0.02
                        data.append(row)
                        
                    matched_layers_major.add(0)

                    layers.append({
                        "upper_depth_major": (major_position, lower_depth_major[idx]),
                        "lower_depth_major": (major_position, lower_depth_major[idx]),
                        "upper_depth_minor": (minor_position, lower_depth_minor[0]),
                        "lower_depth_minor": (minor_position, upper_depth_minor[0]),
                        "label": soil_type_major[idx],
                        "color": color_mapping[str(int(soil_type_minor[idx]))],
                        "soil_type": soil_type_major[idx],
                    })
                                        # 預測predict_borehole_data的數據
                    # 取用layers的數據
                    upper_limit = lower_depth_major[idx] * weight_1 + lower_depth_minor[0] * weight_2
                    lower_limit = lower_depth_major[idx] * weight_1 + upper_depth_minor[0] * weight_2
                    upper_limit = round(upper_limit, 2)
                    lower_limit = round(lower_limit, 2)
                    print(upper_limit, lower_limit)
                    depth = upper_limit
                    if depth - last_depth >= 0.02:
                        print('不改變深度', depth)
                    else:
                        depth = depth + 0.01
                    # 選取data_1和data_2在範圍upper_depth_major、upper_depth_minor、lower_depth_major和lower_depth_minor之間的數據
                    # 使用layers的數據
                    data_major = major_data[(major_data['Depth (m)'] >= upper_limit) & (major_data['Depth (m)'] <= lower_limit)]
                    data_minor = minor_data[(minor_data['Depth (m)'] >= upper_limit) & (minor_data['Depth (m)'] <= lower_limit)]
                    # 初始化變數
                    x = 0

                    # 遍歷深度範圍
                    while depth < lower_limit:
                        # 先檢查索引 x 是否在 data_major 和 data_minor 範圍內
                        row_major = data_major.iloc[x] if x < len(data_major) else None
                        row_minor = data_minor.iloc[x] if x < len(data_minor) else None

                        # 判斷 row_major 和 row_minor 是否有數據，並計算合併或單項數據
                        if row_major is not None and row_minor is not None:
                            # 合併數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': (row_major.iloc[1] * weight_1 + row_minor.iloc[1] * weight_2),
                                'fs (MPa)': (row_major.iloc[2] * weight_1 + row_minor.iloc[2] * weight_2),
                                'u (MPa)': (row_major.iloc[3] * weight_1 + row_minor.iloc[3] * weight_2),
                            }
                        elif row_major is not None and row_minor is None:
                            # 僅使用 row_major 的數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': row_major.iloc[1],
                                'fs (MPa)': row_major.iloc[2],
                                'u (MPa)': row_major.iloc[3],
                            }
                        elif row_minor is not None and row_major is  None:
                            # 僅使用 row_minor 的數據
                            row = {
                                'Depth (m)': depth,
                                'qc (MPa)': row_minor.iloc[1],
                                'fs (MPa)': row_minor.iloc[2],
                                'u (MPa)': row_minor.iloc[3],
                            }
                        # 更新索引和深度
                        x += 1
                        last_depth = depth
                        depth += 0.02
                        data.append(row)
                        
                    matched_layers_minor.add(idx)

        for i in include:
            print('include', i)
            layers.append({
                "upper_depth_major": (major_position, upper_depth_major[idx]),
                "lower_depth_major": (major_position, upper_depth_major[idx]),
                "upper_depth_minor": (minor_position, upper_depth_minor[i]),
                "lower_depth_minor": (minor_position, lower_depth_minor[i]),
                "label": soil_type_minor[i],
                "color": color_mapping[str(int(soil_type_minor[i]))],
                "soil_type": soil_type_minor[i],
            })
            # 預測predict_borehole_data的數據
            # 取用layers的數據
            upper_limit = upper_depth_major[idx] * weight_1 + upper_depth_minor[i] * weight_2
            lower_limit = upper_depth_major[idx] * weight_1 + lower_depth_minor[i] * weight_2
            upper_limit = round(upper_limit, 2)
            lower_limit = round(lower_limit, 2)
            print(upper_limit, lower_limit)
            depth = upper_limit
            if depth - last_depth >= 0.02:
                print('不改變深度', depth)
            else:
                depth = depth + 0.01
            data_major = major_data[(major_data['Depth (m)'] >= upper_limit) & (major_data['Depth (m)'] <= lower_limit)]
            data_minor = minor_data[(minor_data['Depth (m)'] >= upper_limit) & (minor_data['Depth (m)'] <= lower_limit)]
            # 初始化變數
            x = 0

            # 遍歷深度範圍
            while depth < lower_limit:
                # 先檢查索引 x 是否在 data_major 和 data_minor 範圍內
                row_major = data_major.iloc[x] if x < len(data_major) else None
                row_minor = data_minor.iloc[x] if x < len(data_minor) else None

                # 判斷 row_major 和 row_minor 是否有數據，並計算合併或單項數據
                if row_major is not None and row_minor is not None:
                    # 合併數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': (row_major.iloc[1] * weight_1 + row_minor.iloc[1] * weight_2),
                        'fs (MPa)': (row_major.iloc[2] * weight_1 + row_minor.iloc[2] * weight_2),
                        'u (MPa)': (row_major.iloc[3] * weight_1 + row_minor.iloc[3] * weight_2),
                    }
                elif row_major is not None and row_minor is None:
                    # 僅使用 row_major 的數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': row_major.iloc[1],
                        'fs (MPa)': row_major.iloc[2],
                        'u (MPa)': row_major.iloc[3],
                    }
                elif row_minor is not None and row_major is  None:
                    # 僅使用 row_minor 的數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': row_minor.iloc[1],
                        'fs (MPa)': row_minor.iloc[2],
                        'u (MPa)': row_minor.iloc[3],
                    }

                # 更新索引和深度
                x += 1
                depth += 0.02
                data.append(row)
                last_depth = depth
            matched_layers_minor.add(i)
    # 匹配剩下的
    for i in range(len(minor_section)):
        if i not in matched_layers_minor:
            layers.append({
                "upper_depth_major": (major_position, lower_depth_major[idx]),
                "lower_depth_major": (major_position, lower_depth_major[idx]),
                "upper_depth_minor": (minor_position, upper_depth_minor[i]),
                "lower_depth_minor": (minor_position, lower_depth_minor[i]),
                "label": soil_type_minor[i],
                "color": color_mapping[str(int(soil_type_minor[i]))],
                "soil_type": soil_type_minor[i],
            })
                        # 預測predict_borehole_data的數據
            # 取用layers的數據
            upper_limit = lower_depth_major[idx] * weight_1 + upper_depth_minor[i] * weight_2
            lower_limit = lower_depth_major[idx] * weight_1 + lower_depth_minor[i] * weight_2
            upper_limit = round(upper_limit, 2)
            lower_limit = round(lower_limit, 2)
            
            print(upper_limit, lower_limit)
            depth = upper_limit
            if depth - last_depth >= 0.02:
                print('不改變深度', depth)
            else:
                depth = depth + 0.01
            data_major = major_data[(major_data['Depth (m)'] >= upper_limit) & (major_data['Depth (m)'] <= lower_limit)]
            data_minor = minor_data[(minor_data['Depth (m)'] >= upper_limit) & (minor_data['Depth (m)'] <= lower_limit)]
            # 初始化變數
            x = 0
            # 遍歷深度範圍
            while depth < lower_limit:
                # 先檢查索引 x 是否在 data_major 和 data_minor 範圍內
                row_major = data_major.iloc[x] if x < len(data_major) else None
                row_minor = data_minor.iloc[x] if x < len(data_minor) else None

                # 判斷 row_major 和 row_minor 是否有數據，並計算合併或單項數據
                if row_major is not None and row_minor is not None:
                    # 合併數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': (row_major.iloc[1] * weight_1 + row_minor.iloc[1] * weight_2),
                        'fs (MPa)': (row_major.iloc[2] * weight_1 + row_minor.iloc[2] * weight_2),
                        'u (MPa)': (row_major.iloc[3] * weight_1 + row_minor.iloc[3] * weight_2),
                    }
                elif row_major is not None and row_minor is None:
                    # 僅使用 row_major 的數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': row_major.iloc[1],
                        'fs (MPa)': row_major.iloc[2],
                        'u (MPa)': row_major.iloc[3],
                    }
                elif row_minor is not None and row_major is  None:
                    # 僅使用 row_minor 的數據
                    row = {
                        'Depth (m)': depth,
                        'qc (MPa)': row_minor.iloc[1],
                        'fs (MPa)': row_minor.iloc[2],
                        'u (MPa)': row_minor.iloc[3],
                    }

                # 更新索引和深度
                x += 1
                last_depth = depth
                depth += 0.02
                data.append(row)
                
            matched_layers_minor.add(idx)
    # 最後一次性轉換為 DataFrame
    predict_borehole_data = pd.DataFrame(data)
# 把predict_borehole_data的順序改為Depth (m)由小到大
predict_borehole_data = predict_borehole_data.sort_values(by='Depth (m)', ascending=True)
# 刪掉重複的數據
predict_borehole_data = predict_borehole_data.drop_duplicates(subset='Depth (m)', keep='first')



# 把 layers 的資料轉換成 DataFrame
df_layers = pd.DataFrame(layers)



predict_borehole['Type'] = df_layers['soil_type']

# 分別取出深度值進行計算，並四捨五入到小數點後兩位
predict_borehole['Upper Depth'] = (
    df_layers['upper_depth_major'].apply(lambda x: x[1]) * weight_1 + 
    df_layers['upper_depth_minor'].apply(lambda x: x[1]) * weight_2
).round(2)

predict_borehole['Lower Depth'] = (
    df_layers['lower_depth_major'].apply(lambda x: x[1]) * weight_1 + 
    df_layers['lower_depth_minor'].apply(lambda x: x[1]) * weight_2
).round(2)



# 儲存預測的鑽孔位置
predict_borehole.to_excel('predict_borehole.xlsx', index=False)

# 儲存預測的鑽孔資料
predict_borehole_data.to_excel('predict_borehole_data.xlsx', index=False)




# 繪圖
fig, ax = plt.subplots(figsize=(12, 8))

# 建立不重複的圖例
used_labels = set()
legend_handles = []

for layer in layers:
    upper_depth_major = layer["upper_depth_major"]
    lower_depth_major = layer["lower_depth_major"]
    upper_depth_minor = layer["upper_depth_minor"]
    lower_depth_minor = layer["lower_depth_minor"]
    points = [upper_depth_major, lower_depth_major, lower_depth_minor, upper_depth_minor] # 更新 points 為四點的列表
    label = layer["label"]
    color = layer["color"]

    # 定義多邊形，使用四個點構成的列表
    polygon = Polygon(points, closed=True, color=color, alpha=0.7)
    ax.add_patch(polygon)
    
    if label not in used_labels:
        legend_handles.append(plt.Rectangle((0, 0), 1, 1, fc=color, alpha=0.7, label=label))
        used_labels.add(label)

# 添加鑽孔位置線
ax.axvline(x=borehole_position_1, color='black', linestyle='--', linewidth=1, label='Borehole 1')
ax.axvline(x=borehole_position_2, color='black', linestyle='--', linewidth=1, label='Borehole 2')
ax.axvline(x=780, color='black', linestyle='--', linewidth=1, label='Borehole 2')
ax.axvline(x=785, color='black', linestyle='--', linewidth=1, label='Borehole 2')

# 設置圖例
ax.legend(handles=legend_handles, loc='upper right', bbox_to_anchor=(1.15, 1))

# 設置軸和標題
ax.yaxis.set_major_locator(ticker.MultipleLocator(10))
plt.gca().invert_yaxis()
ax.set_xlim(0, borehole_position_2)
ax.set_ylim(105, 0)
ax.set_title("Soil Type Visualization between Boreholes")
ax.set_xlabel("Distance (m)")
ax.set_ylabel("Depth (m)")

plt.tight_layout()
# 儲存圖像
plt.savefig('soil_type_visualization.png', dpi=300)
plt.show()