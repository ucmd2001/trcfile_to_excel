import pandas as pd
import os

file_path = input("請輸入 .trc 文件的完整路徑: ")

# 替换路径中的空格为转义字符（如果有必要）
file_path = file_path.replace(' ', '\\ ')

df = pd.read_csv(
     file_path,
     delim_whitespace=True,
     header=None,
     skiprows=16,
)
# 篩選出 E 欄包含 '0481', '0482', '0483', '0484', '0485' 的行
values_to_find = ['0481', '0482', '0483', '0484', '0485']
filtered_df = df[df[3].isin(values_to_find)].copy() # 建立一個副本

# 在 filtered_df 中建立一個新列，包含前一行的第 10 列的值
filtered_df['previous_value'] = filtered_df.iloc[:, 10].shift(1)

for i in filtered_df.index:
    if filtered_df.loc[i, 'previous_value'] == '11' and filtered_df.loc[i, 10] == '01':
        target_value = filtered_df.loc[i, 3]
        break

# 如果找到了符合條件的行
if 'target_value' in locals():
    # 再次篩選，只包含第 3 列值與 target_value 相同的行
    final_filtered_df = filtered_df[filtered_df[3] == target_value]
    final_filtered_df[10] = final_filtered_df[10].apply(lambda x: 11 if x == '11' else x)
    # 找到符合條件的行在 final_filtered_df 中的位置
    position = final_filtered_df.index.tolist().index(i)

    # 提取所需行
    selected_rows = final_filtered_df.index[:]
    selected_data = final_filtered_df.loc[selected_rows]
    # 获取文件名（不含扩展名）
    base_name = os.path.basename(file_path)  # 获取文件的基本名称
    file_name_without_extension = os.path.splitext(base_name)[0]  # 去掉扩展名

    # 使用相同的文件名创建一个新的 Excel 文件
    output_file_name = file_name_without_extension + '.xlsx'
    selected_data.to_excel(output_file_name, index=False)

else:
     print("沒有找到符合條件的行")