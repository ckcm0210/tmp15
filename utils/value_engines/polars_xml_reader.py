import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Optional

# 以 XML 解析 .xlsx 的 worksheet 值（cached），再交由上層使用（可配合 Polars 做後處理）
# 返回結構：{ sheet_name: { 'A1': value, ... } }

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def _load_shared_strings(z: zipfile.ZipFile) -> list:
    sst = []
    try:
        if 'xl/sharedStrings.xml' not in z.namelist():
            return sst
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.findall(f'{{{NS_MAIN}}}si'):
            # 可能有多個 t
            text_parts = []
            # 先找 r/t（富文本）
            for t in si.findall(f'.//{{{NS_MAIN}}}t'):
                text_parts.append(t.text or '')
            if not text_parts:
                # 退回 si/t
                t = si.find(f'{{{NS_MAIN}}}t')
                if t is not None:
                    text_parts.append(t.text or '')
            sst.append(''.join(text_parts))
    except Exception:
        pass
    return sst


def _workbook_sheet_names(z: zipfile.ZipFile) -> list:
    names = []
    try:
        root = ET.fromstring(z.read('xl/workbook.xml'))
        for s in root.findall(f'.//{{{NS_MAIN}}}sheet'):
            nm = s.attrib.get('name')
            if nm:
                names.append(nm)
    except Exception:
        pass
    return names


def _col_letters_to_index(s: str) -> int:
    # 例如 'A'->1, 'Z'->26, 'AA'->27
    s = s.upper()
    n = 0
    for ch in s:
        if 'A' <= ch <= 'Z':
            n = n * 26 + (ord(ch) - ord('A') + 1)
        else:
            break
    return n


def _split_addr(addr: str):
    # 將 'A10' 拆成 ('A', 10)
    col = ''
    row = ''
    for ch in addr:
        if ch.isalpha():
            col += ch
        else:
            row += ch
    try:
        r = int(row)
    except Exception:
        r = 0
    return col, r


def read_values_from_xlsx_via_polars_xml(xlsx_path: str) -> Dict[str, Dict[str, Optional[str]]]:
    out: Dict[str, Dict[str, Optional[str]]] = {}
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            sst = _load_shared_strings(z)
            sheets = _workbook_sheet_names(z)
            # 順序依 workbook.xml
            for i, name in enumerate(sheets, start=1):
                sheet_path = f'xl/worksheets/sheet{i}.xml'
                if sheet_path not in z.namelist():
                    continue
                try:
                    xml = z.read(sheet_path)
                    root = ET.fromstring(xml)
                    vals: Dict[str, Optional[str]] = {}
                    for c in root.findall(f'.//{{{NS_MAIN}}}c'):
                        addr = c.attrib.get('r')
                        if not addr:
                            continue
                        t = c.attrib.get('t')  # s=sharedString, b=boolean, str=string, inlineStr, etc.
                        v_node = c.find(f'{{{NS_MAIN}}}v')
                        if v_node is None:
                            # inlineStr 支援
                            is_node = c.find(f'{{{NS_MAIN}}}is')
                            if is_node is not None:
                                tnode = is_node.find(f'.//{{{NS_MAIN}}}t')
                                vals[addr] = (tnode.text if tnode is not None else '')
                            continue
                        raw = v_node.text
                        if raw is None:
                            vals[addr] = None
                        elif t == 's':
                            # shared string
                            try:
                                idx = int(raw)
                                vals[addr] = sst[idx] if 0 <= idx < len(sst) else ''
                            except Exception:
                                vals[addr] = ''
                        elif t == 'b':
                            vals[addr] = True if raw in ('1', 'true', 'TRUE') else False
                        else:
                            # 數值或一般字串，先原樣返回（上層如需再做型別轉換）
                            vals[addr] = raw
                    out[name] = vals
                except Exception as e:
                    # 單張表失敗不影響其他表
                    out[name] = {}
    except Exception:
        return {}
    return out
