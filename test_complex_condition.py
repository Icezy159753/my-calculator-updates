#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Debug script สำหรับตรวจสอบเงื่อนไข (q030a_r1!=0 | q030a_r3!=0) & Q040a_R3=0
"""

import sys
import os
import importlib.util

# Import ฟังก์ชันที่ต้องการทดสอบ
spec = importlib.util.spec_from_file_location(
    "clean_spss",
    os.path.join(os.path.dirname(__file__), 'All_Programs', '99_CleanSPSS_Germini.py')
)
clean_spss = importlib.util.module_from_spec(spec)
spec.loader.exec_module(clean_spss)

expand_wildcard = clean_spss.expand_wildcard
auto_convert = clean_spss.auto_convert
validate_condition = clean_spss.validate_condition

import pandas as pd

def test_complex_condition():
    """ทดสอบเงื่อนไข (q030a_r1!=0 | q030a_r3!=0) & Q040a_R3=0"""
    
    print("=" * 70)
    print("Debug: เงื่อนไข (q030a_r1!=0 | q030a_r3!=0) & Q040a_R3=0")
    print("=" * 70)
    
    # สร้าง test data
    test_cols = ['q030a_r1', 'q030a_r3', 'Q040a_R3', 'ID']
    lower_to_orig_map = {col.lower(): col for col in test_cols}
    
    # สร้าง dataframe ทดสอบ
    df = pd.DataFrame({
        'q030a_r1': [1, 0, 2, 0, 1],
        'q030a_r3': [0, 1, 0, 0, 2],
        'Q040a_R3': [0, 0, 0, 1, 0],
        'ID': [1, 2, 3, 4, 5]
    })
    
    print("\nTest Data:")
    print(df)
    print()
    
    condition = "(q030a_r1!=0 | q030a_r3!=0) & Q040a_R3=0"
    
    print(f"Original Condition: {condition}\n")
    
    try:
        # Validate
        error = validate_condition(condition, test_cols, set(c.lower() for c in test_cols), lower_to_orig_map)
        if error:
            print(f"❌ Validation Error: {error}")
            return
        print("✓ Validation: OK")
        
        # Expand wildcard
        expanded = expand_wildcard(condition, test_cols, lower_to_orig_map)
        print(f"\nExpanded: {expanded}")
        
        # Convert
        converted = auto_convert(expanded, lower_to_orig_map)
        print(f"Converted: {converted}")
        
        # Evaluate
        print(f"\nEvaluating with df.eval()...")
        result = df.eval(converted)
        count = int(result.sum())
        matched_ids = df[result]['ID'].tolist() if result.any() else []
        
        print(f"✓ Count: {count}")
        print(f"Matched IDs: {matched_ids}")
        
        print("\nExpected Result:")
        print("  ID 1: q030a_r1=1(!=0) | q030a_r3=0 → True | False = True, Q040a_R3=0 → True & True = TRUE ✓")
        print("  ID 2: q030a_r1=0 | q030a_r3=1(!=0) → False | True = True, Q040a_R3=0 → True & True = TRUE ✓")
        print("  ID 3: q030a_r1=2(!=0) | q030a_r3=0 → True | False = True, Q040a_R3=0 → True & True = TRUE ✓")
        print("  ID 4: q030a_r1=0 | q030a_r3=0 → False | False = False, Q040a_R3=1 → False & False = FALSE")
        print("  ID 5: q030a_r1=1(!=0) | q030a_r3=2(!=0) → True | True = True, Q040a_R3=0 → True & True = TRUE ✓")
        print("\nExpected Matched IDs: [1, 2, 3, 5]")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
    
    print("\n" + "=" * 70)

if __name__ == '__main__':
    test_complex_condition()
