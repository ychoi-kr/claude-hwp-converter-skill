#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HWPX to TXT Converter (v0.1.0)

Extracts text content from Korean HWPX (Hangul Word Processor XML) files.
"""

import sys
import os
import zipfile
import xml.etree.ElementTree as ET


def extract_text_from_xml(xml_path):
    """
    Extract text from XML file with multiple namespace attempts.
    
    Args:
        xml_path: Path to the XML file
        
    Returns:
        Extracted text content as string
    """
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # Try multiple namespaces
        namespaces = [
            {'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph'},
            {'hp': 'http://www.hancom.co.kr/hwpml/2010/main'},
            {'hp': 'http://www.hancom.co.kr/hwpml/2016/paragraph'}
        ]
        
        for namespace in namespaces:
            try:
                text_parts = []
                for para in root.findall(".//hp:p", namespace):
                    line = []
                    for run in para.findall(".//hp:run", namespace):
                        text_elem = run.find("hp:t", namespace)
                        if text_elem is not None and text_elem.text:
                            line.append(text_elem.text)
                    if line:
                        text_parts.append(''.join(line))
                if text_parts:
                    return '\n'.join(text_parts)
            except:
                continue
        
        # Try without namespace
        try:
            text_parts = []
            for para in root.findall(".//p"):
                for run in para.findall(".//run"):
                    for text_elem in run.findall(".//t"):
                        if text_elem.text:
                            text_parts.append(text_elem.text)
            if text_parts:
                return '\n'.join(text_parts)
        except:
            pass
        
        return ""
    except Exception as e:
        return f"[Error]: {e}"


def extract_from_prvtext(hwpx_path):
    """
    Fallback method: extract from PrvText.txt inside HWPX.
    
    Args:
        hwpx_path: Path to the HWPX file
        
    Returns:
        Extracted text content as string
    """
    try:
        with zipfile.ZipFile(hwpx_path, 'r') as zip_ref:
            for file_info in zip_ref.infolist():
                if "PrvText.txt" in file_info.filename:
                    with zip_ref.open(file_info) as file:
                        return file.read().decode('utf-8', errors='ignore')
    except Exception as e:
        return f"[Error] Could not extract from PrvText: {e}"
    return ""


def extract_text_from_hwpx(hwpx_path):
    """
    Extract text from HWPX file.
    
    Args:
        hwpx_path: Path to the HWPX file
        
    Returns:
        Extracted text content as string
    """
    extracted_dir = "/tmp/hwpx_extracted"
    os.makedirs(extracted_dir, exist_ok=True)
    
    # Extract HWPX
    try:
        with zipfile.ZipFile(hwpx_path, 'r') as zip_ref:
            zip_ref.extractall(extracted_dir)
    except Exception as e:
        return f"[Error] Extraction failed: {e}"
    
    # Process XML files
    full_text = ""
    contents_dir = os.path.join(extracted_dir, "Contents")
    
    if os.path.isdir(contents_dir):
        xml_files = [f for f in os.listdir(contents_dir) if f.endswith('.xml')]
        for file_name in sorted(xml_files):
            xml_path = os.path.join(contents_dir, file_name)
            extracted_content = extract_text_from_xml(xml_path)
            if extracted_content and not extracted_content.startswith("[Error"):
                full_text += extracted_content + "\n"
    
    # PrvText.txt fallback if XML extraction failed
    if not full_text.strip():
        print("Warning: XML extraction incomplete. Using PrvText.txt fallback...")
        full_text = extract_from_prvtext(hwpx_path)
    
    return full_text


def convert_hwpx(input_path, output_path="/tmp/converted_text.txt"):
    """
    Convert HWPX file to text.
    
    Args:
        input_path: Path to HWPX file
        output_path: Path for output text file
        
    Returns:
        bool: True if conversion succeeded, False otherwise
    """
    if not os.path.exists(input_path):
        print(f"❌ File not found: {input_path}")
        return False
    
    if not input_path.endswith('.hwpx'):
        print(f"❌ Not a HWPX file: {input_path}")
        return False
    
    print("🔧 Extracting text from XML sections...")
    
    # Extract text
    extracted_text = extract_text_from_hwpx(input_path)
    
    # Check if extracted text is empty
    if not extracted_text.strip():
        print("❌ No text could be extracted from the file")
        return False
    
    # Save to output location
    try:
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(extracted_text.strip())
    except Exception as e:
        print(f"❌ Failed to save output: {str(e)}")
        return False
    
    print(f"✅ Conversion complete! Output saved to: {output_path}")
    print(f"📊 Text length: {len(extracted_text):,} characters")
    print(f"📊 Lines: {extracted_text.count(chr(10)) + 1:,}")
    
    # Preview (first 1000 characters)
    preview_length = min(1000, len(extracted_text))
    if preview_length > 0:
        print("\n" + "=" * 60)
        print("📖 Preview (first 1000 characters):")
        print("=" * 60)
        print(extracted_text[:preview_length])
        if len(extracted_text) > preview_length:
            print("\n... (truncated)")
    
    return True


def main():
    """Command-line interface entry point"""
    if len(sys.argv) < 2:
        print("Usage: python -m scripts.hwpx.converter <file.hwpx> [output.txt]")
        sys.exit(1)
    
    hwpx_file = sys.argv[1]
    
    if len(sys.argv) >= 3:
        output_path = sys.argv[2]
    else:
        output_path = "/tmp/converted_text.txt"
    
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    
    success = convert_hwpx(hwpx_file, output_path)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
