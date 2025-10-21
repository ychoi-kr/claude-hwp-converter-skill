#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
HWP 5.x Record Parser

Handles:
- Text cleaning and control character processing
- Record header parsing
- Table structure parsing
- Section content extraction
"""

import struct
import zlib
import re
from typing import List, Tuple, Optional
from dataclasses import dataclass


# ============================================================
# 🧱 HWP Tag Constants
# ============================================================

HWPTAG_BEGIN = 0x10
HWPTAG_PARA_HEADER = HWPTAG_BEGIN + 50  # 0x42
HWPTAG_PARA_TEXT = HWPTAG_BEGIN + 51  # 0x43
HWPTAG_LIST_HEADER = HWPTAG_BEGIN + 56  # 0x48
HWPTAG_TABLE = HWPTAG_BEGIN + 61  # 0x4D


# ============================================================
# 🎮 Control Character Detection (based on pyhwp)
# ============================================================

# 제어 문자 크기 정의 (WCHAR 단위, not bytes!)
CONTROL_CHAR_SIZES = {
    # CHAR 타입: 1 WCHAR
    0x00: 1,   # NULL
    0x0a: 1,   # LINE_BREAK
    0x0d: 1,   # PARAGRAPH_BREAK
    0x18: 1,   # HYPHEN
    0x1e: 1,   # NONBREAK_SPACE
    0x1f: 1,   # FIXWIDTH_SPACE
    
    # INLINE/EXTENDED 타입: 8 WCHARs
    0x01: 8,   # EXTENDED
    0x02: 8,   # SECTION_COLUMN_DEF
    0x03: 8,   # FIELD_START
    0x04: 8,   # FIELD_END (INLINE)
    0x05: 8,   # INLINE
    0x06: 8,   # INLINE
    0x07: 8,   # INLINE
    0x08: 8,   # TITLE_MARK
    0x09: 8,   # TAB
    0x0b: 8,   # DRAWING_TABLE_OBJECT
    0x0c: 8,   # EXTENDED
    0x0e: 8,   # EXTENDED
    0x0f: 8,   # HIDDEN_EXPLANATION
    0x10: 8,   # HEADER_FOOTER
    0x11: 8,   # FOOT_END_NOTE
    0x12: 8,   # AUTO_NUMBER
    0x13: 8,   # INLINE
    0x14: 8,   # INLINE
    0x15: 8,   # PAGE_CTLCHR
    0x16: 8,   # BOOKMARK
    0x17: 8,   # EXTENDED
}


def find_control_char(data: bytes, start: int) -> Tuple[int, int]:
    """
    Find the next control character in UTF-16LE data.
    Based on pyhwp's implementation with regex.
    
    Returns:
        (control_start, control_end) positions in bytes
    """
    # UTF-16LE 제어 문자 패턴: [\x00-\x1f]\x00
    import re
    REGEX_CONTROL_CHAR = re.compile(b'[\x00-\x1f]\x00')
    
    while True:
        m = REGEX_CONTROL_CHAR.search(data, start)
        if m is None:
            return len(data), len(data)
        
        i = m.start()
        
        # 홀수 인덱스면 잘못된 매칭 (UTF-16LE는 2바이트 정렬)
        if i & 1 == 1:
            start = i + 1
            continue
        
        # 제어 문자 코드
        ch = data[i]
        
        # 제어 문자 크기 (WCHAR 단위)
        if ch in CONTROL_CHAR_SIZES:
            size = CONTROL_CHAR_SIZES[ch]
            return i, i + (size * 2)  # WCHAR → bytes
        else:
            # 알 수 없는 제어 문자는 1 WCHAR로 간주
            return i, i + 2


def parse_para_text_chunks(data: bytes) -> List[str]:
    """
    Parse PARA_TEXT and extract text chunks, skipping control chars.
    Based on pyhwp's implementation.
    
    Args:
        data: Raw PARA_TEXT bytes (UTF-16LE)
        
    Returns:
        List of text strings
    """
    texts = []
    idx = 0
    size = len(data)
    
    while idx < size:
        ctrlpos, ctrlpos_end = find_control_char(data, idx)
        
        # 제어 문자 이전의 텍스트 추출
        if idx < ctrlpos:
            try:
                text = data[idx:ctrlpos].decode('utf-16le', errors='ignore')
                if text:
                    texts.append(text)
            except Exception:
                pass
        
        # 다음 위치로 이동
        idx = ctrlpos_end if ctrlpos_end > idx else idx + 2
    
    return texts


# ============================================================
# 🧹 Control Character Cleaning
# ============================================================

CONTROL_PLACEHOLDERS = {
    0x0B: "",              # 그리기/표 앵커 → 제거 (표는 별도 파싱)
    0x10: "",              # 머리말/꼬리말 경계 → 제거
    0x11: "",              # 각주/미주 → 제거 (본문만 추출)
    0x15: "\n",            # 페이지 컨트롤 → 줄바꿈
    0x18: "-",             # 하이픈 유지
    0x1E: " ",             # 묶음 빈칸
    0x1F: " ",             # 고정폭 빈칸
}
KEEP = {9, 10, 13}         # 탭, LF, CR


def clean_hwp_text(text: str, in_table: bool = False) -> str:
    """
    Clean HWP text by handling control characters and special unicode.
    
    Args:
        text: Raw text from HWP file
        in_table: If True, apply table-specific cleaning
        
    Returns:
        Cleaned text with control characters replaced or removed
    """
    # 1) 제어 코드 치환/유지/제거
    out_chars = []
    for ch in text:
        code = ord(ch)
        if code >= 32:
            out_chars.append(ch)
        elif code in KEEP:
            # 표 안에서는 줄바꿈을 공백으로 변환
            if in_table and code in (10, 13):
                out_chars.append(" ")
            else:
                out_chars.append(ch)
        elif code in CONTROL_PLACEHOLDERS:
            placeholder = CONTROL_PLACEHOLDERS[code]
            if placeholder:  # 빈 문자열이 아닌 경우만 추가
                out_chars.append(placeholder)
        # else: 제거 (정보 없는 제어문자)

    text = "".join(out_chars)

    # 2) 특수 유니코드 정리
    text = (text
        .replace("\ufeff", "")  # BOM
        .replace("\u200b", "")  # Zero-width space
        .replace("\u200c", "")  # Zero-width non-joiner
        .replace("\u200d", "")  # Zero-width joiner
    )

    # 3) 공백/개행 정리
    if in_table:
        # 표 안: 여러 공백을 하나로, 개행 제거
        text = re.sub(r"[ \t]+", " ", text)
        text = text.replace("\n", " ").replace("\r", " ")
        text = re.sub(r"\s+", " ", text)
    else:
        # 일반 텍스트: 공백 정리, 과도한 개행 제거
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
    
    return text.strip()


# ============================================================
# 🧩 Data Classes
# ============================================================

@dataclass
class TableCell:
    """Represents a single cell in HWP table"""
    col: int
    row: int
    col_span: int
    row_span: int
    text: str


@dataclass
class Table:
    """Represents an HWP table with cells"""
    row_count: int
    col_count: int
    cells: List[TableCell]

    def to_text(self) -> str:
        """
        Convert table to formatted text with box-drawing characters.
        Handles cell merging (colspan/rowspan).
        
        Returns:
            Formatted table as string
        """
        if not self.cells:
            return ""
        
        # 그리드 생성 - 병합 정보를 고려
        grid = [[None for _ in range(self.col_count)] for _ in range(self.row_count)]
        
        for cell in self.cells:
            if cell.row >= self.row_count or cell.col >= self.col_count:
                continue
            
            # 병합된 셀 처리
            for r in range(cell.row, min(cell.row + cell.row_span, self.row_count)):
                for c in range(cell.col, min(cell.col + cell.col_span, self.col_count)):
                    if r == cell.row and c == cell.col:
                        # 병합의 시작 셀에만 텍스트 저장
                        grid[r][c] = cell.text
                    else:
                        # 병합된 나머지 셀은 빈 문자열
                        if grid[r][c] is None:
                            grid[r][c] = ""
        
        # None을 빈 문자열로 변환
        for r in range(self.row_count):
            for c in range(self.col_count):
                if grid[r][c] is None:
                    grid[r][c] = ""
        
        # 열 너비 계산 (병합 고려)
        col_widths = [3] * self.col_count  # 최소 너비 3
        for cell in self.cells:
            if cell.col >= self.col_count:
                continue
            text_len = len(cell.text)
            if cell.col_span == 1:
                col_widths[cell.col] = max(col_widths[cell.col], text_len)
            else:
                # 병합된 셀의 경우 너비를 분산
                avg_width = text_len // cell.col_span
                for c in range(cell.col, min(cell.col + cell.col_span, self.col_count)):
                    col_widths[c] = max(col_widths[c], avg_width)
        
        # 테이블 그리기
        lines = []
        lines.append("┌" + "┬".join("─" * (w + 2) for w in col_widths) + "┐")
        
        for r in range(self.row_count):
            row_texts = []
            for c in range(self.col_count):
                cell_text = grid[r][c]
                row_texts.append(cell_text.ljust(col_widths[c]))
            lines.append("│ " + " │ ".join(row_texts) + " │")
            
            if r < self.row_count - 1:
                lines.append("├" + "┼".join("─" * (w + 2) for w in col_widths) + "┤")
        
        lines.append("└" + "┴".join("─" * (w + 2) for w in col_widths) + "┘")
        return "\n".join(lines)


# ============================================================
# 🧠 Core Parsing Functions
# ============================================================

def read_hwp_metadata(ole) -> dict:
    """
    Extract version and compression info from FileHeader.
    
    Args:
        ole: OleReader instance
        
    Returns:
        Dictionary with version and compressed flag
    """
    if "FileHeader" not in ole.list_streams():
        return {"version": None, "compressed": None}
    
    data = ole.read_stream("FileHeader")
    version_bytes = data[32:36]
    version = struct.unpack("<I", version_bytes)[0]
    major, minor, build, revision = (
        (version >> 24) & 0xFF,
        (version >> 16) & 0xFF,
        (version >> 8) & 0xFF,
        version & 0xFF,
    )
    flags = struct.unpack("<I", data[36:40])[0] if len(data) >= 40 else 0
    compressed = bool(flags & 0x01)
    return {"version": f"{major}.{minor}.{build}.{revision}", "compressed": compressed}


def read_record_header(data: bytes, pos: int) -> Tuple[Optional[int], Optional[int], Optional[int], int]:
    """
    Parse HWP record header.
    
    Args:
        data: Binary data
        pos: Current position
        
    Returns:
        Tuple of (tag_id, level, size, new_position)
    """
    if pos + 4 > len(data):
        return None, None, None, pos
    
    header = struct.unpack('<I', data[pos:pos+4])[0]
    tag_id = header & 0x3FF
    level = (header >> 10) & 0x3FF
    size = (header >> 20) & 0xFFF
    pos += 4
    
    if size == 0xFFF:
        if pos + 4 > len(data):
            return None, None, None, pos
        size = struct.unpack('<I', data[pos:pos+4])[0]
        pos += 4
    
    return tag_id, level, size, pos


def parse_table_header(data: bytes, pos: int) -> Tuple[int, int, int]:
    """
    Parse table header to get row and column count.
    
    Args:
        data: Binary data
        pos: Current position
        
    Returns:
        Tuple of (row_count, col_count, new_position)
    """
    if pos + 8 > len(data):
        return 0, 0, pos
    
    pos += 4
    row_count = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    col_count = struct.unpack('<H', data[pos:pos+2])[0]
    pos += 2
    return row_count, col_count, pos


def parse_cell_list(data: bytes, pos: int, end: int, level: int) -> List[TableCell]:
    """
    Parse table cells from binary data.
    
    Args:
        data: Binary data
        pos: Start position
        end: End position
        level: Current record level
        
    Returns:
        List of TableCell objects
    """
    cells = []
    while pos < end:
        tag_id, rec_level, size, new_pos = read_record_header(data, pos)
        if tag_id is None or new_pos + size > len(data):
            break
        
        if tag_id == HWPTAG_LIST_HEADER and rec_level == level + 1:
            cell_pos = new_pos
            if cell_pos + 26 <= len(data):
                col = struct.unpack('<H', data[cell_pos:cell_pos+2])[0]
                row = struct.unpack('<H', data[cell_pos+2:cell_pos+4])[0]
                col_span = struct.unpack('<H', data[cell_pos+4:cell_pos+6])[0]
                row_span = struct.unpack('<H', data[cell_pos+6:cell_pos+8])[0]
                
                text_parts = []
                cell_data_pos = cell_pos + size
                cell_end_pos = new_pos + size
                
                while cell_data_pos < cell_end_pos:
                    para_tag, _, para_size, para_pos = read_record_header(data, cell_data_pos)
                    if para_tag is None or para_pos + para_size > len(data):
                        break
                    
                    if para_tag == HWPTAG_PARA_TEXT:
                        text_data = data[para_pos:para_pos+para_size]
                        # pyhwp 방식: 제어 문자 처리
                        text_chunks = parse_para_text_chunks(text_data)
                        for text in text_chunks:
                            text = clean_hwp_text(text, in_table=True)
                            if text:
                                text_parts.append(text)
                    
                    cell_data_pos = para_pos + para_size
                
                text = " ".join(text_parts)
                cells.append(TableCell(col, row, col_span, row_span, text))
        
        pos = new_pos + size
    
    return cells


def extract_content_from_section(ole, section_name: str) -> Tuple[List[str], List[Table]]:
    """
    Extract paragraphs and tables from a section.
    
    Args:
        ole: OleReader instance
        section_name: Name of the section stream (e.g., "BodyText/Section0")
        
    Returns:
        Tuple of (paragraphs list, tables list)
    """
    if section_name not in ole.list_streams():
        return [], []
    
    section_data = ole.read_stream(section_name)
    original_size = len(section_data)
    decompressed = False
    
    # 압축 해제 시도
    for wbits in [-15, 15, 0]:
        try:
            section_data = zlib.decompress(section_data, wbits)
            decompressed = True
            print(f"   ✓ Decompressed: {original_size} → {len(section_data)} bytes (zlib {wbits})")
            break
        except Exception:
            continue
    
    if not decompressed:
        print(f"   ! Not compressed or unknown format: {len(section_data)} bytes")

    paragraphs, tables = [], []
    pos = 0
    
    # 문단 헤더를 추적하여 텍스트 길이 확인
    current_para_nchars = None
    
    while pos < len(section_data):
        tag_id, level, size, new_pos = read_record_header(section_data, pos)
        if tag_id is None or new_pos + size > len(section_data):
            break
        
        # PARA_HEADER 파싱 (HWPTAG_PARA_HEADER = 0x43-1 = 0x42)
        if tag_id == HWPTAG_BEGIN + 50:  # PARA_HEADER
            if new_pos + 4 <= len(section_data):
                # nchars 읽기 (첫 4바이트)
                nchars = struct.unpack('<I', section_data[new_pos:new_pos+4])[0]
                # MSB가 1이면 제거
                if nchars & 0x80000000:
                    nchars &= 0x7FFFFFFF
                current_para_nchars = nchars
        
        elif tag_id == HWPTAG_PARA_TEXT:
            text_data = section_data[new_pos:new_pos+size]
            
            # pyhwp 방식: 제어 문자를 찾아서 텍스트 청크 추출
            text_chunks = parse_para_text_chunks(text_data)
            for text in text_chunks:
                text = clean_hwp_text(text)
                if text:
                    paragraphs.append(text)
            
            # 다음 문단을 위해 리셋
            current_para_nchars = None
        
        elif tag_id == HWPTAG_TABLE:
            rows, cols, tpos = parse_table_header(section_data, new_pos)
            if rows > 0 and cols > 0:
                cells = parse_cell_list(section_data, tpos, new_pos + size, level)
                if cells:
                    tables.append(Table(rows, cols, cells))
        
        pos = new_pos + size
    
    return paragraphs, tables
