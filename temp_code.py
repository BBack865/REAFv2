# 두번째 문장이 "NACL" 또는 "Dil"인 경우 특수 처리
if len(next_parts) > 1 and (next_parts[1] == "NACL" or next_parts[1] == "Dil"):
    # "-"가 포함된 단어를 찾아 AU로 설정
    au = ""
    for word in next_parts:
        if "-" in word:
            au = word
            break
    
    # "-"가 포함된 단어를 찾지 못한 경우 기본 방식으로 처리
    if not au and next_parts[1] == "NACL" and len(next_parts) > 2:
        au = next_parts[2]
    
    rp_lot = ""
    if len(next_parts) >= 6:  # 다섯번째 문장 (인덱스 4)
        potential_rp_lot = next_parts[4]
        # 숫자로만 구성되어 있거나 숫자가 포함된 문자열인 경우만 R.P Lot으로 인식
        if potential_rp_lot.isdigit() or (potential_rp_lot.isalnum() and any(c.isdigit() for c in potential_rp_lot)):
            rp_lot = potential_rp_lot
else:
    # 일반적인 경우: "-"가 포함된 단어를 찾아 AU로 설정
    au = ""
    for word in next_parts:
        if "-" in word:
            au = word
            break
    
    # "-"가 포함된 단어를 찾지 못한 경우 기본 방식으로 처리
    if not au and len(next_parts) > 1:
        au = next_parts[1]
    
    rp_lot = ""
    if len(next_parts) >= 5 and len(next_parts) > 3:
        potential_rp_lot = next_parts[3]
        # 숫자로만 구성되어 있거나 숫자가 포함된 문자열인 경우만 R.P Lot으로 인식
        if potential_rp_lot.isdigit() or (potential_rp_lot.isalnum() and any(c.isdigit() for c in potential_rp_lot)):
            rp_lot = potential_rp_lot
