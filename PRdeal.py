import pandas as pd
import os

# --- 配置 ---
excel_file_name = '0827.xls' 
sheet_names = ['Sheet1', 'Sheet2', 'Sheet3']
# --- 修改点 1: 将输出文件名改为 .xlsx ---
output_filename = 'extracted_data.xlsx'

# --- 数据处理 ---
all_data_frames = []

print("开始处理 Excel 文件...")

if not os.path.exists(excel_file_name):
    print(f"错误：文件 '{excel_file_name}' 不存在。请确保它和脚本在同一个文件夹中。")
else:
    for sheet in sheet_names:
        print(f"正在读取文件: '{excel_file_name}' 的工作表: '{sheet}'")
        
        try:
            df = pd.read_excel(excel_file_name, sheet_name=sheet, header=0, skiprows=[1])
        except Exception as e:
            print(f"读取工作表 '{sheet}' 时出错: {e}")
            continue

        columns_to_extract = [
            col for col in df.columns 
            if isinstance(col, str) and ('荷重' in col or '位移' in col)
        ]
        
        if not columns_to_extract:
            print(f"在工作表 '{sheet}' 中没有找到'荷重'或'位移'列。")
            continue
            
        extracted_df = df[columns_to_extract].copy()
        extracted_df.dropna(how='all', inplace=True)
        
        rename_dict = {}
        load_count = 1
        disp_count = 1
        for col in extracted_df.columns:
            if '荷重' in col:
                rename_dict[col] = f'{sheet}_荷重_{load_count}'
                load_count += 1
            elif '位移' in col:
                rename_dict[col] = f'{sheet}_位移_{disp_count}'
                disp_count += 1
        extracted_df.rename(columns=rename_dict, inplace=True)

        all_data_frames.append(extracted_df)

if all_data_frames:
    combined_df = pd.concat([df.reset_index(drop=True) for df in all_data_frames], axis=1)
    
    # --- 修改点 2: 使用 to_excel() 保存文件 ---
    # index=False 表示在输出的 Excel 文件中不包含行号
    combined_df.to_excel(output_filename, index=False)
    
    print("\n-----------------------------------------")
    print(f"✅ 数据提取完成！")
    print(f"所有'荷重'和'位移'的数据已合并并保存到 Excel 文件: '{output_filename}'")
    print("-----------------------------------------")
else:
    print("\n处理完成，但没有提取到任何数据。")