import os
import re

def find_target_files(input_folder):
    """파일명 중간의 '_숫자_C' 패턴을 찾습니다."""
    all_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".mpt")]
    temp_file_map = {}
    extracted_temps = []
    pattern = re.compile(r'_(\d{3,4})C?_C') 

    for f in all_files:
        match = pattern.search(f)
        if match:
            t_str = match.group(1).strip()
            if t_str.endswith(('00', '50')):
                if t_str not in extracted_temps:
                    extracted_temps.append(t_str)
                    temp_file_map[t_str] = f
                
    return temp_file_map, extracted_temps

def get_data_start_line(file_path):
    try:
        with open(file_path, 'r', encoding='cp1252') as f:
            for i, line in enumerate(f):
                if "freq/Hz" in line: return i
    except: return 65