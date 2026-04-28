import pandas as pd
import os
import re

def main():
    # 使用局部变量替代全局变量
    result_dist = {}
    
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(current_dir, '统计表.xlsx')
    print(f"读取文件: {file_path}")

    try:
        df = pd.read_excel(file_path)
        print(f"文件读取成功，共 {len(df)} 行数据")
        
        for index, row in df.iterrows():
            row_data = {}
            for i, value in enumerate(row[6:], 1):
                # 假设需要正则提取的列包含特定字符，或者明确记录第20列是什么数据
                # 如果某一天问卷格式变了，这里需要相应修改
                if i == 20 and isinstance(value, str) and "〖" in value:
                    pattern = r"〖(.*?)〗"
                    matches = re.findall(pattern, value)
                    if matches:
                        value = matches[0]
                
                row_data[f'data{i}'] = value
            result_dist[index] = row_data
            
        return result_dist
    except Exception as e:
        print(f"读取文件时出错: {e}")
        return {} # 返回空字典而不是None，防止下游调用时报错

def print_dist(dist):
    print("读取完成！")
    print(f"总共有 {len(dist)} 行数据")
    print("前5行数据示例：")
    
    for i, (key, value) in enumerate(list(dist.items())[:5]):
        print(f"行 {key+2}: {value}")

if __name__ == '__main__':
    data_dict = main()
    if data_dict:
        print_dist(data_dict)