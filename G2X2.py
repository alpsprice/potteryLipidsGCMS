import os
import glob
import re
import pandas as pd

# version 3.1
#This file was first uploaded on Aug 15, 2025 by LHX from IVPP; Under the MIT License;
# ===== User Configuration =====
# Working folder (default: script location)
INPUT_FOLDER = os.path.dirname(os.path.realpath(__file__))
# Output Excel file name
OUTPUT_FILE = "GCMS_Summary.xlsx"
# Category columns (you can add more categories)
CATEGORIES = ['SFA', 'UFA', 'DC', 'Plant', 'Animal']

# ---------- Library switch ----------
# False: use embedded RULES below; True: use external Excel library named lib2026.xlsx placed in the same folder as this script;
USE_EXTERNAL_LIB = False
EXTERNAL_LIB_FILE = "lib2026.xlsx"

# ---------- Embedded rules (used when USE_EXTERNAL_LIB = False) ----------
# Each rule contains: name(list), category, value, si_threshold, ratio_expected, ratio_tolerance
# 分析规则（可扩展）, ratio_expected 指[目标的保留时间减 C16:0 的]除以 [C18:0 和 C16:0 的时间差], 符合则大概率可信并自动标*号; 名称可写多个，可用正则表达式
RULES = [
    #名称前后加^和$是为了锁死全文匹配, 比如C10和C16名称可局部匹配. 问题来源是为了用正则式通配一些化合物, 求教更优方案ing
    #C9:0
    {
        'name': ['^Nonanoic acid, methyl ester$'],
        'category': 'SFA',
        'value': 'C9:0',
        'si_threshold': 85,
        'ratio_expected': -4.2,
        'ratio_tolerance': 0.09
    },
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
    #C26:0 27及更长链基本没可能自动识别出来, 自己检查
    {
        'name': ['Hexacosanoic acid, methyl ester'],
        'category': 'SFA',
        'value': 'C26:0',
        'si_threshold': 80,
        'ratio_expected': 4.5,
        'ratio_tolerance': 0.08
    },
    #C16:1
    {
        'name': [r'Methyl hexadec-9-enoate'],
        'category': 'UFA',
        'value': 'C16:1',
        'si_threshold': 80,
        'ratio_expected': -0.2,
        'ratio_tolerance': 0.9
    },
    #C18:1
    {
        'name': [r'9-Octadecenoic acid.* methyl ester.*'],
        'category': 'UFA',
        'value': 'C18:1',
        'si_threshold': 80,
        'ratio_expected': 0.9,
        'ratio_tolerance': 0.1
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
        'category': 'DC',
        'value': 'C4',
        'si_threshold': 80,
        'ratio_expected': -5.66,
        'ratio_tolerance': 0.07
    },
    #C5
    {
        'name': ['^Pentanedioic acid, dimethyl ester$'],
        'category': 'DC',
        'value': 'C5',
        'si_threshold': 80,
        'ratio_expected': -4.8,
        'ratio_tolerance': 0.09
    },
    #C6
    {
        'name': ['^Hexanedioic acid, dimethyl ester$'],
        'category': 'DC',
        'value': 'C6',
        'si_threshold': 71,
        'ratio_expected': -4,
        'ratio_tolerance': 0.07
    },
    #C7
    {
        'name': [r'Hexanedioic acid .*-methyl.* dimethyl ester', 'Heptanedioic acid, dimethyl ester'],
        'category': 'DC',
        'value': 'C7',
        'si_threshold': 80,
        'ratio_expected': -3.38,
        'ratio_tolerance': 0.06
    },
    #C8
    {
        'name': ['^Octanedioic acid, dimethyl ester$'],
        'category': 'DC',
        'value': 'C8',
        'si_threshold': 85,
        'ratio_expected': -2.63,
        'ratio_tolerance': 0.06
    },
    #C9
    {
        'name': ['^Nonanedioic acid, dimethyl ester$'],
        'category': 'DC',
        'value': 'C9',
        'si_threshold': 85,
        'ratio_expected': -2.1,
        'ratio_tolerance': 0.1
    },
    #C10
    {
        'name': ['^Decanedioic acid, dimethyl ester$'],
        'category': 'DC',
        'value': 'C10',
        'si_threshold': 85,
        'ratio_expected': -1.5,
        'ratio_tolerance': 0.05
    },
    #C11
    {
        'name': ['^Undecanedioic acid, dimethyl ester$'],
        'category': 'DC',
        'value': 'C11',
        'si_threshold': 85,
        'ratio_expected': -0.93,
        'ratio_tolerance': 0.05
    },
    #C12
    {
        'name': ['^Dodecanedioic acid, dimethyl ester$'],
        'category': 'DC',
        'value': 'C12',
        'si_threshold': 85,
        'ratio_expected': -0.39,
        'ratio_tolerance': 0.05
    },
    #C13
    {
        'name': ['^Tridecanedioic acid, dimethyl ester$'],
        'category': 'DC',
        'value': 'C13',
        'si_threshold': 85,
        'ratio_expected': 0.14,
        'ratio_tolerance': 0.2
    },
    #C14 这个自动基本识别不到
    {
        'name': ['^Dimethyl tetradecanedioate$', '^Tetradecanedioic acid, dimethyl ester$'],
        'category': 'DC',
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
    # b-Sitosterol
    {
        'name': ['.beta.-Sitosterol acetate'],
        'category': 'Plant',
        'value': 'b-Sitosterol acetate',
        'si_threshold': 70,
        'ratio_expected': 4.5,
        'ratio_tolerance': 0.04
    },
]

# ---------- Peak area ratio calculation rules (extendable) ----------
# Each rule is a dictionary with:
#   value1 : value of the first compound (must match a 'value' in RULES)
#   mode1  : 'both' (default, ±tolerance), 'gt' (greater than lower bound), or 'lt' (less than upper bound)
#   value2 : value of the second compound
#   mode2  : same as mode1
#   output : output column name
RATIO_CALCULATIONS = [
    # C18:1 / C18:0
    {
        'value1': 'C18:1',
        'mode1': 'both',
        'value2': 'C18:0',
        'mode2': 'both',
        'output': 'C18:1/C18:0'
    },
    # A/P
    {
        'value1': 'C9',
        'mode1': 'both',
        'value2': 'C16:0',
        'mode2': 'both',
        'output': 'A/P'
    },
    # Add more rules as needed, e.g.:
    # {
    #     'value1': 'C16:0',
    #     'mode1': 'both',
    #     'value2': 'C18:0',
    #     'mode2': 'both',
    #     'output': 'C16:0/C18:0'
    # },
]

# ===== End of Configuration =====

def load_external_rules(filepath):
    """Load rules from external Excel file, return list of rules (same format as RULES)"""
    if not os.path.exists(filepath):
        print(f"Warning: external library file {filepath} not found. Falling back to embedded rules.")
        return None
    try:
        df = pd.read_excel(filepath)
        required_cols = ['name', 'category', 'value', 'si_threshold', 'ratio_expected', 'ratio_tolerance']
        for col in required_cols:
            if col not in df.columns:
                print(f"Error: external Excel missing required column '{col}'")
                return None
        rules = []
        for _, row in df.iterrows():
            name_str = str(row['name']).strip()
            if not name_str:
                continue
            # Split by semicolon and strip
            names = [n.strip() for n in name_str.split(';') if n.strip()]
            if not names:
                continue
            try:
                rule = {
                    'name': names,
                    'category': str(row['category']).strip(),
                    'value': str(row['value']).strip(),
                    'si_threshold': int(row['si_threshold']),
                    'ratio_expected': float(row['ratio_expected']),
                    'ratio_tolerance': float(row['ratio_tolerance'])
                }
                rules.append(rule)
            except (ValueError, TypeError) as e:
                print(f"Skipping invalid row: {row.to_dict()}, error: {e}")
        print(f"Successfully loaded {len(rules)} rules from external library.")
        return rules
    except Exception as e:
        print(f"Failed to read external library: {e}. Falling back to embedded rules.")
        return None

def name_matches(peak_name, patterns):
    """Check if peak_name matches any pattern in patterns (regex or substring)"""
    for pattern in patterns:
        try:
            if re.search(pattern, peak_name, re.IGNORECASE):
                return True
        except re.error:
            if pattern.lower() in peak_name.lower():
                return True
    return False

def meets_ratio_condition(ratio, expected, tolerance, mode='both'):
    """
    Check if ratio meets condition based on mode:
    - 'both': |ratio - expected| <= expected * tolerance
    - 'gt'  : ratio >= expected - expected * tolerance
    - 'lt'  : ratio <= expected + expected * tolerance
    """
    if expected is None or tolerance is None:
        return True  # no condition
    lower = expected - abs(expected) * tolerance
    upper = expected + abs(expected) * tolerance
    if mode == 'both':
        return lower <= ratio <= upper
    elif mode == 'gt':
        return ratio >= lower
    elif mode == 'lt':
        return ratio <= upper
    else:
        return True  # unknown mode, treat as no condition

def parse_data_block(block, rules, value_to_rule, ratio_calcs):
    """Parse a single data block and return dictionary with results"""
    result = {'Name': '', 'RefTimeDiff': None}
    for cat in CATEGORIES:
        result[cat] = []
    for calc in ratio_calcs:
        result[calc['output']] = None

    # Extract file name
    file_match = re.search(r"Data File Name\t(.+?\.qgd)", block)
    if file_match:
        file_path = file_match.group(1)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        result['Name'] = file_name

    # Extract peak table
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
        # Ensure area field exists (try common variants)
        if 'Area' not in peak_data:
            for key in peak_data:
                if key.lower() == 'area':
                    peak_data['Area'] = peak_data[key]
                    break
            else:
                peak_data['Area'] = None
        peaks.append(peak_data)

    # Find reference compounds (C16:0 and C18:0)
    hexadecanoic = None
    stearate = None
    for peak in peaks:
        name = peak.get('Name', '').strip()
        if name == 'Hexadecanoic acid, methyl ester':
            try:
                hexadecanoic = {
                    'rt': float(peak['Ret.Time']),
                    'si': int(peak.get('SI', 0))
                }
            except (ValueError, TypeError):
                pass
        elif name == 'Methyl stearate':
            try:
                stearate = {
                    'rt': float(peak['Ret.Time']),
                    'si': int(peak.get('SI', 0))
                }
            except (ValueError, TypeError):
                pass

    # Calculate reference time difference
    ref_time_diff = None
    if hexadecanoic and stearate:
        ref_time_diff = stearate['rt'] - hexadecanoic['rt']
        result['RefTimeDiff'] = ref_time_diff

    # Apply classification rules (take peak with largest area for each rule)
    for rule in rules:
        candidates = []
        for peak in peaks:
            peak_name = peak.get('Name', '').strip()
            if not name_matches(peak_name, rule['name']):
                continue
            try:
                si = int(peak.get('SI', 0))
                if si < rule['si_threshold']:
                    continue
                area = float(peak.get('Area', 0)) if peak.get('Area') not in (None, '') else 0
                rt = float(peak['Ret.Time'])
            except (ValueError, TypeError):
                continue
            candidates.append((area, rt, peak))
        if not candidates:
            continue
        # Select candidate with largest area
        best = max(candidates, key=lambda x: x[0])
        area, rt, peak = best
        value = rule['value']
        # Check ratio condition for asterisk
        if ref_time_diff and hexadecanoic:
            time_diff = rt - hexadecanoic['rt']
            ratio = time_diff / ref_time_diff
            expected = rule['ratio_expected']
            tolerance = rule['ratio_tolerance']
            if meets_ratio_condition(ratio, expected, tolerance, mode='both'):
                value += '*'
        result[rule['category']].append(value)

    # Apply ratio calculations
    if ref_time_diff and hexadecanoic:
        for calc in ratio_calcs:
            # Process first compound
            rule1 = value_to_rule.get(calc['value1'])
            if not rule1:
                continue
            candidates1 = []
            for peak in peaks:
                peak_name = peak.get('Name', '').strip()
                if not name_matches(peak_name, rule1['name']):
                    continue
                try:
                    si = int(peak.get('SI', 0))
                    if si < rule1['si_threshold']:
                        continue
                    area = float(peak.get('Area', 0)) if peak.get('Area') not in (None, '') else 0
                    rt = float(peak['Ret.Time'])
                except (ValueError, TypeError):
                    continue
                time_diff = rt - hexadecanoic['rt']
                ratio = time_diff / ref_time_diff
                # Apply mode-specific condition
                mode = calc.get('mode1', 'both')
                if not meets_ratio_condition(ratio, rule1['ratio_expected'], rule1['ratio_tolerance'], mode):
                    continue
                candidates1.append((area, ratio))
            area1 = max([c[0] for c in candidates1]) if candidates1 else None

            # Process second compound
            rule2 = value_to_rule.get(calc['value2'])
            if not rule2:
                continue
            candidates2 = []
            for peak in peaks:
                peak_name = peak.get('Name', '').strip()
                if not name_matches(peak_name, rule2['name']):
                    continue
                try:
                    si = int(peak.get('SI', 0))
                    if si < rule2['si_threshold']:
                        continue
                    area = float(peak.get('Area', 0)) if peak.get('Area') not in (None, '') else 0
                    rt = float(peak['Ret.Time'])
                except (ValueError, TypeError):
                    continue
                time_diff = rt - hexadecanoic['rt']
                ratio = time_diff / ref_time_diff
                mode = calc.get('mode2', 'both')
                if not meets_ratio_condition(ratio, rule2['ratio_expected'], rule2['ratio_tolerance'], mode):
                    continue
                candidates2.append((area, ratio))
            area2 = max([c[0] for c in candidates2]) if candidates2 else None

            if area1 is not None and area2 is not None and area2 != 0:
                result[calc['output']] = area1 / area2
            else:
                result[calc['output']] = None
    else:
        # No reference, cannot calculate ratios
        for calc in ratio_calcs:
            result[calc['output']] = None

    return result

def main():
    # Determine which rule set to use
    if USE_EXTERNAL_LIB:
        ext_file = os.path.join(INPUT_FOLDER, EXTERNAL_LIB_FILE)
        rules = load_external_rules(ext_file)
        if rules is None:
            print("Using embedded rules as fallback.")
            rules = RULES
        else:
            print("Using external library rules.")
    else:
        rules = RULES
        print("Using embedded rules.")

    # Build mapping from value to rule (assuming each value is unique)
    value_to_rule = {rule['value']: rule for rule in rules if 'value' in rule}

    all_results = []

    # Process all .txt files in input folder
    for file_path in glob.glob(os.path.join(INPUT_FOLDER, "*.txt")):
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Split into blocks by [Header]
        blocks = re.split(r"\[Header\]", content)[1:]  # skip first empty element

        for block in blocks:
            result = parse_data_block(block, rules, value_to_rule, RATIO_CALCULATIONS)
            # Convert lists to comma-separated strings
            for key in result:
                if isinstance(result[key], list):
                    result[key] = ", ".join(result[key])
            all_results.append(result)

    # Create DataFrame
    df = pd.DataFrame(all_results)

    # Ensure column order: Name + categories + ratio columns
    base_columns = ['Name'] + CATEGORIES
    ratio_columns = [calc['output'] for calc in RATIO_CALCULATIONS]
    all_columns = base_columns + ratio_columns
    for col in all_columns:
        if col not in df.columns:
            df[col] = ''

    # Save to Excel
    df[all_columns].to_excel(OUTPUT_FILE, index=False)
    print(f"Processing completed! {len(all_results)} data blocks processed.")
    print(f"Results saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()