import os
import glob
import re
import pandas as pd
from openpyxl import Workbook

#This file was first uploaded on Aug 15, 2025 by LHX from IVPP; Under the MIT License;
# ===== user config =====
# Input your working folder (default folder is where you are placing this script)
INPUT_FOLDER = os.path.dirname(os.path.realpath(__file__)) #也可以用绝对路径 r"D:\GCMS\BLL"
# 输出Excel文件名
OUTPUT_FILE = "GCMS_Summary.xlsx"
# 定义类目列（可扩展）
CATEGORIES = ['SFA', 'UFA', 'DI', 'Plant', 'Animal']
# 定义分析规则（可扩展）, ratio_expected 指[目标的保留时间减 C16:0 的]除以 [C18:0 和 C16:0 的时间差], 符合则大概率可信并自动标*号; 名称可写多个，可用正则表达式
RULES = [
    #名称前\后加^$是为了锁死全文匹配, 比如C10和C16名称可局部匹配. 问题来源是为了用正则式通配一些化合物, 求教更优方案ing
    #C10:0
    {
        'name': ['^Decanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C10:0',
        'si_threshold': 85,
        'ratio_expected': -3.5,
        'ratio_tolerance': 0.09
    },
    #C11:0
    {
        'name': ['^Undecanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C11:0',
        'si_threshold': 85,
        'ratio_expected': -2.89,
        'ratio_tolerance': 0.08
    },
    #C12:0
    {
        'name': ['^Dodecanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C12:0',
        'si_threshold': 85,
        'ratio_expected': -2.23,
        'ratio_tolerance': 0.07
    },
    #C13:0
    {
        'name': ['^Tridecanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C13:0',
        'si_threshold': 85,
        'ratio_expected': -1.62,
        'ratio_tolerance': 0.05
    },
    #C14:0
    {
        'name': ['^Methyl tetradecanoate$', '^Myristic acid, methyl ester$', r'Tridecanoic acid.*, methyl ester'],
        'category': 'SFA',
        'value': 'C14:0',
        'si_threshold': 85,
        'ratio_expected': -1.08,
        'ratio_tolerance': 0.05
    },
    #C15:0
    {
        'name': ['^Pentadecanoic acid, methyl ester$', '^Methyl pentadecanoate$'],
        'category': 'SFA',
        'value': 'C15:0',
        'si_threshold': 85,
        'ratio_expected': -0.525,
        'ratio_tolerance': 0.06
    },
    #C16:0
    {
        'name': ['^Hexadecanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C16:0',
        'si_threshold': 85,
        'ratio_expected': 0,
        'ratio_tolerance': 0.02
    },
    #C17:0
    {
        'name': ['^Heptadecanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C17:0',
        'si_threshold': 85,
        'ratio_expected': 0.51,
        'ratio_tolerance': 0.03
    },
    #C18:0
    {
        'name': ['^Methyl stearate$'],
        'category': 'SFA',
        'value': 'C18:0',
        'si_threshold': 85,
        'ratio_expected': 1,
        'ratio_tolerance': 0.01
    },
    #C19:0
    {
        'name': ['^Nonadecanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C19:0',
        'si_threshold': 85,
        'ratio_expected': 1.5,
        'ratio_tolerance': 0.05
    },
    #C20:0
    {
        'name': ['^Arachidic acid, methyl ester$', '^Eicosanoic acid, methyl ester$', '^Methyl arachisate$', r'Methyl .*-meth.*nonadecanoate'], 
        'category': 'SFA',
        'value': 'C20:0',
        'si_threshold': 85,
        'ratio_expected': 2.1,
        'ratio_tolerance': 0.1
    },
    #C21:0
    {
        'name': ['^Heneicosanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C21:0',
        'si_threshold': 71,
        'ratio_expected': 2.6,
        'ratio_tolerance': 0.08
    },
    #C22:0
    {
        'name': ['^Docosanoic acid, methyl ester$', r'Methyl .*-meth.*heneicosanoate'],
        'category': 'SFA',
        'value': 'C22:0',
        'si_threshold': 85,
        'ratio_expected': 2.97,
        'ratio_tolerance': 0.07
    },
    #C23:0
    {
        'name': ['^Tricosanoic acid, methyl ester$', 'Methyl tricosanoate'],
        'category': 'SFA',
        'value': 'C23:0',
        'si_threshold': 85,
        'ratio_expected': 3.4,
        'ratio_tolerance': 0.05
    },
    #C24:0
    {
        'name': ['^Tetracosanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C24:0',
        'si_threshold': 85,
        'ratio_expected': 3.75,
        'ratio_tolerance': 0.08
    },
    #C25:0
    {
        'name': ['^Pentacosanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C25:0',
        'si_threshold': 71,
        'ratio_expected': 4.15,
        'ratio_tolerance': 0.08
    },
    #C26:0 27及更长链基本没可能自动识别出来，自己检查
    {
        'name': ['Hexacosanoic acid, methyl ester'],
        'category': 'SFA',
        'value': 'C26:0',
        'si_threshold': 80,
        'ratio_expected': 4.5,
        'ratio_tolerance': 0.08
    },
    #C18:1
    {
        'name': [r'9-Octadecenoic acid.* methyl ester.*'],
        'category': 'UFA',
        'value': 'C18:1',
        'si_threshold': 82,
        'ratio_expected': 0.9,
        'ratio_tolerance': 0.06
    },
    #C20:1
    {
        'name': [r'cis-Methyl .*-eicosenoate', r'.*-Eicosenoic acid, methyl ester.*'],
        'category': 'UFA',
        'value': 'C20:1',
        'si_threshold': 80,
        'ratio_expected': 1.95,
        'ratio_tolerance': 0.08
    },
    #C22:1
    {
        'name': ['13-Docosenoic acid, methyl ester', r'Methyl .*-docosenoate'],
        'category': 'UFA',
        'value': 'C22:1',
        'si_threshold': 85,
        'ratio_expected': 2.80,
        'ratio_tolerance': 0.07
    },
    #C18:2
    {
        'name': [r'Methyl .*-trans,.*-cis-octadecadienoate'],
        'category': 'UFA',
        'value': 'C18:2',
        'si_threshold': 85,
        'ratio_expected': 0.834,
        'ratio_tolerance': 0.06
    },
    #C4  DI类如果有，可以理论出现位置，或者在 SFA 后0.3 s内排查，人工确定具体有哪些
    {
        'name': ['^Butanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C4',
        'si_threshold': 80,
        'ratio_expected': -5.66,
        'ratio_tolerance': 0.07
    },
    #C5
    {
        'name': ['^Pentanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C5',
        'si_threshold': 80,
        'ratio_expected': -4.8,
        'ratio_tolerance': 0.09
    },
    #C6
    {
        'name': ['^Hexanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C6',
        'si_threshold': 71,
        'ratio_expected': -4,
        'ratio_tolerance': 0.07
    },
    #C7
    {
        'name': [r'Hexanedioic acid .*-methyl.* dimethyl ester', 'Heptanedioic acid, dimethyl ester'],
        'category': 'DI',
        'value': 'C7',
        'si_threshold': 80,
        'ratio_expected': -3.38,
        'ratio_tolerance': 0.06
    },
    #C8
    {
        'name': ['^Octanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C8',
        'si_threshold': 85,
        'ratio_expected': -2.63,
        'ratio_tolerance': 0.06
    },
    #C9
    {
        'name': ['^Nonanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C9',
        'si_threshold': 85,
        'ratio_expected': -2.1,
        'ratio_tolerance': 0.05
    },
    #C10
    {
        'name': ['^Decanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C10',
        'si_threshold': 85,
        'ratio_expected': -1.5,
        'ratio_tolerance': 0.05
    },
    #C11
    {
        'name': ['^Undecanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C11',
        'si_threshold': 85,
        'ratio_expected': -0.93,
        'ratio_tolerance': 0.05
    },
    #C12
    {
        'name': ['^Dodecanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C12',
        'si_threshold': 85,
        'ratio_expected': -0.39,
        'ratio_tolerance': 0.05
    },
    #C13
    {
        'name': ['^Tridecanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C13',
        'si_threshold': 85,
        'ratio_expected': 0.14,
        'ratio_tolerance': 0.2
    },
    #C14 这个自动基本识别不到
    {
        'name': ['^Dimethyl tetradecanedioate$', '^Tetradecanedioic acid, dimethyl ester$'],
        'category': 'DI',
        'value': 'C14',
        'si_threshold': 75,
        'ratio_expected': 0.635,
        'ratio_tolerance': 0.06
    },
    #Cholestanol #未定
    {
        'name': ['Cholestanol'],
        'category': 'Animal',
        'value': 'Cholestanol',
        'si_threshold': 85,
        'ratio_expected': 4.09, #未定
        'ratio_tolerance': 0.04
    },
    #Ergostanol
    {
        'name': ['Ergostanol'],
        'category': 'Plant',
        'value': 'Ergostanol',
        'si_threshold': 85,
        'ratio_expected': 4.18,
        'ratio_tolerance': 0.04
    },
    #b-谷甾醇
    {
        'name': ['.beta.-Sitosterol acetate'],
        'category': 'Plant',
        'value': 'b-Sitosterol acetate',
        'si_threshold': 70,
        'ratio_expected': 4.5,
        'ratio_tolerance': 0.04
    },
    # 添加更多规则示例：
    # {
    #     'name': ['Hexadecane'],
    #     'category': 'ALK',
    #     'value': 'C16',
    #     'si_threshold': 80,
    #     'ratio_expected': 1.0,
    #     'ratio_tolerance': 0.1
    # },
]
# ===== 配置结束 =====

def parse_data_block(block):
    """解析单个数据块"""
    result = {'名称': '', '基准时间差': None}
    for cat in CATEGORIES:
        result[cat] = []
    
    # 提取文件名
    file_match = re.search(r"Data File Name\t(.+?\.qgd)", block)
    if file_match:
        file_path = file_match.group(1)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        result['名称'] = file_name
    
    # 提取峰表数据
    peak_table_match = re.search(r"\[MC Peak Table\][\s\S]+?(?=\[Header\]|$)", block)
    if not peak_table_match:
        return result
    
    peak_table = peak_table_match.group(0)
    header_match = re.search(r"Peak#\tRet\.Time\t.*?Name\t.*?SI\t", peak_table)
    if not header_match:
        return result
    
    header_line = header_match.group(0)
    headers = [h.strip() for h in header_line.split('\t')]
    
    peaks = []
    for line in peak_table.splitlines():
        if line.startswith('Peak#') or not line.strip() or 'Header' in line:
            continue
        values = line.split('\t')
        if len(values) < len(headers):
            continue
        peak_data = dict(zip(headers, values))
        peaks.append(peak_data)
    
    # 查找基准化合物
    hexadecanoic = None
    stearate = None
    
    # 基准化合物名称列表
    hexadecanoic_names = ['Hexadecanoic acid, methyl ester']
    stearate_names = ['Methyl stearate']
    
    for peak in peaks:
        name = peak.get('Name', '').strip()
        if name in hexadecanoic_names:
            try:
                hexadecanoic = {
                    'rt': float(peak['Ret.Time']),
                    'si': int(peak.get('SI', 0))
                }
            except (ValueError, TypeError):
                pass
        elif name in stearate_names:
            try:
                stearate = {
                    'rt': float(peak['Ret.Time']),
                    'si': int(peak.get('SI', 0))
                }
            except (ValueError, TypeError):
                pass
    
    # 计算基准时间差
    if hexadecanoic and stearate:
        result['基准时间差'] = stearate['rt'] - hexadecanoic['rt']
    
    # 应用规则
    for rule in RULES:
        matched = False
        for peak in peaks:
            peak_name = peak.get('Name', '').strip()
            
            # 检查峰名称是否匹配规则中的任一模式（支持正则表达式）
            name_match = False
            for pattern in rule['name']:
                try:
                    # 尝试正则表达式匹配
                    if re.search(pattern, peak_name, re.IGNORECASE):
                        name_match = True
                        break
                except re.error:
                    # 如果正则表达式无效，尝试直接字符串匹配
                    if pattern.lower() in peak_name.lower():
                        name_match = True
                        break
            
            if not name_match:
                continue
            
            try:
                si = int(peak.get('SI', 0))
                rt = float(peak['Ret.Time'])
            except (ValueError, TypeError):
                continue
            
            if si < rule['si_threshold']:
                continue
            
            value = rule['value']
            # 检查保留时间比例
            if result['基准时间差'] and hexadecanoic:
                time_diff = rt - hexadecanoic['rt']
                ratio = time_diff / result['基准时间差']
                expected = rule['ratio_expected']
                tolerance = rule['ratio_tolerance']
                
                if abs(ratio - expected) <= abs(expected * tolerance):
                    value += '*'
            
            result[rule['category']].append(value)
            matched = True
            break  # 只取第一个匹配项
    
    return result

def main():
    all_results = []
    
    # 处理所有文本文件
    for file_path in glob.glob(os.path.join(INPUT_FOLDER, "*.txt")):
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 分割数据块
        blocks = re.split(r"\[Header\]", content)[1:]  # 跳过第一个空元素
        
        for block in blocks:
            result = parse_data_block(block)
            # 将列表转换为逗号分隔的字符串
            for key in result:
                if isinstance(result[key], list):
                    result[key] = ", ".join(result[key])
            all_results.append(result)
    
    # 创建DataFrame
    df = pd.DataFrame(all_results)
    
    # 确保列顺序正确
    columns = ['名称'] + CATEGORIES
    for col in columns:
        if col not in df.columns:
            df[col] = ''
    
    # 保存Excel
    df[columns].to_excel(OUTPUT_FILE, index=False)
    print(f"处理完成! 共处理 {len(all_results)} 个数据块")
    print(f"结果已保存至: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
