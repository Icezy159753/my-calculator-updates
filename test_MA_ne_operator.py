#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script เพื่อทดสอบปัญหา MA != operator ว่าได้รับการแก้ไขแล้ว
"""

import sys
import os
import importlib.util

# Import ฟังก์ชันที่ต้องการทดสอบ
# Note: The module name has leading numbers, so we import it differently
import importlib.util
spec = importlib.util.spec_from_file_location(
    "clean_spss",
    os.path.join(os.path.dirname(__file__), 'All_Programs', '99_CleanSPSS_Germini.py')
)
clean_spss = importlib.util.module_from_spec(spec)
spec.loader.exec_module(clean_spss)

expand_wildcard = clean_spss.expand_wildcard
auto_convert = clean_spss.auto_convert
validate_condition = clean_spss.validate_condition
STRICT_MODE = clean_spss.STRICT_MODE
import pandas as pd

def test_ma_ne_operator():
    """ทดสอบ != operator กับ MA questions"""
    
    print("=" * 60)
    print("ทดสอบ MA != Operator Fix")
    print("=" * 60)
    
    # สร้าง test data
    test_cols = ['Q023_O1', 'Q023_O2', 'Q023_O3', 'ID']
    lower_to_orig_map = {col.lower(): col for col in test_cols}
    
    # สร้าง dataframe ทดสอบ
    df = pd.DataFrame({
        'Q023_O1': [1, 2, 3, 1, None],
        'Q023_O2': [2, 1, 1, 2, ''],
        'Q023_O3': [3, 3, 1, 3, None],
        'ID': [1, 2, 3, 4, 5]
    })
    
    print("\nTest Data:")
    print(df)
    print()
    
    # Test cases
    test_cases = [
        ('Q023_O!=1', 'ทุก Q023_O columns ต้อง != 1'),
        ('Q023_O=1', 'อย่างน้อยหนึ่ง Q023_O column ต้อง = 1'),
        ('Q023_O!=1,2', 'ทุก Q023_O columns ต้อง != 1,2'),
    ]
    
    for condition, description in test_cases:
        print(f"\nTest Case: {condition}")
        print(f"Description: {description}")
        
        try:
            # Validate
            error = validate_condition(condition, test_cols, set(c.lower() for c in test_cols), lower_to_orig_map)
            if error:
                print(f"  ❌ Validation Error: {error}")
                continue
            
            # Expand wildcard
            expanded = expand_wildcard(condition, test_cols, lower_to_orig_map)
            print(f"  Expanded: {expanded}")
            
            # Convert
            converted = auto_convert(expanded, lower_to_orig_map)
            print(f"  Converted: {converted}")
            
            # Evaluate
            result = df.eval(converted)
            count = int(result.sum())
            matched_ids = df[result]['ID'].tolist() if result.any() else []
            
            print(f"  ✓ Count: {count}")
            print(f"  Matched IDs: {matched_ids}")
            
        except Exception as e:
            print(f"  ❌ Error: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n" + "=" * 60)
    print("Test Complete")
    print("=" * 60)

if __name__ == '__main__':
    test_ma_ne_operator()
