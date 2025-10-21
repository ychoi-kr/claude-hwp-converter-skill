#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HWP 5.x to TXT Converter (v0.1.0)

Entry point for HWP binary format conversion.
"""

import os
import sys
from .ole import OleReader
from .parser import read_hwp_metadata, extract_content_from_section


def extract_full_text_from_hwp(hwp_path):
    """
    Extract full text and tables from HWP file.
    
    Args:
        hwp_path: Path to HWP file
        
    Returns:
        Extracted text as string, or error message starting with "[Error]"
    """
    if not os.path.exists(hwp_path):
        return f"[Error] File not found: {hwp_path}"
    
    try:
        with OleReader(hwp_path) as ole:
            meta = read_hwp_metadata(ole)
            print(f"Version: {meta['version']}, Compressed: {meta['compressed']}")
            
            all_paragraphs, all_tables = [], []
            idx = 0
            
            # 모든 섹션 순회
            while True:
                name = f"BodyText/Section{idx}"
                if name not in ole.list_streams():
                    break
                
                print(f"📄 Processing {name}...")
                paras, tbls = extract_content_from_section(ole, name)
                all_paragraphs.extend(paras)
                all_tables.extend(tbls)
                idx += 1
            
            # 결과 조합
            result = []
            result.extend(all_paragraphs)
            
            for i, t in enumerate(all_tables, 1):
                result.append(f"\n\n[표 {i}]")
                result.append(t.to_text())
            
            return "\n\n".join(result)
    
    except Exception as e:
        return f"[Error] Failed: {e}"


def convert_hwp(input_path, output_path="/tmp/converted_text.txt"):
    """
    Convert HWP file to plain text.
    
    Args:
        input_path: Path to input HWP file
        output_path: Path to output text file
        
    Returns:
        True if conversion succeeded, False otherwise
    """
    if not os.path.exists(input_path):
        print(f"❌ File not found: {input_path}")
        return False
    
    print("🔧 Extracting full text and tables...")
    text = extract_full_text_from_hwp(input_path)
    
    if text.startswith("[Error]") or not text.strip():
        print(text)
        return False
    
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(text.strip())
        print(f"✅ Output saved: {output_path}")
        return True
    except Exception as e:
        print(f"❌ Save failed: {e}")
        return False


def main():
    """Command-line interface entry point"""
    if len(sys.argv) < 2:
        print("Usage: python -m scripts.hwp5.converter <file.hwp> [output.txt]")
        sys.exit(1)
    
    hwp_file = sys.argv[1]
    
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        output_path = "/tmp/converted_text.txt"
    
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    success = convert_hwp(hwp_file, output_path)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
