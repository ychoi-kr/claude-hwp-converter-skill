---
name: hwp-converter
description: Convert Hancom Word Processor documents (.hwp and .hwpx) to plain UTF-8 text. Use this skill when a user uploads or requests conversion of HWP/HWPX files.
---

# HWP / HWPX Converter

Extracts full readable text and tables from Korean Hangul Word Processor documents.

## Instructions

Run the HWP5 converter for `.hwp` files (HWP 5.x binary format):
```bash
cd /mnt/skills/user/hwp-converter
python -m scripts.hwp5.converter <file_path> [output_path]
```

Run the HWPX converter for `.hwpx` files (HWPX XML format):
```bash
cd /mnt/skills/user/hwp-converter
python -m scripts.hwpx.converter <file_path> [output_path]
```

## Examples

### Converting `.hwp` file

```bash
cd /mnt/skills/user/hwp-converter
python -m scripts.hwp5.converter /mnt/user-data/uploads/document.hwp /mnt/user-data/outputs/document.txt
```

### Converting `.hwpx` file

```bash
cd /mnt/skills/user/hwp-converter
python -m scripts.hwpx.converter /mnt/user-data/uploads/document.hwpx /mnt/user-data/outputs/document.txt
```

### Reading the output

After successful execution, read the output file:
```bash
cat /mnt/user-data/outputs/document.txt
```
