# -*- coding: utf-8 -*-
"""Customer BOM to HQ BOM standardization preview."""

import re

from shared import _cell_str, _open_workbook
from manufacturer_alias import lookup_manufacturer


def split_tokens(value, split_spaces=True):
    text = _cell_str(value)
    if not text:
        return []
    pattern = r'[,;\s\u3001\uff0c\uff1b]+' if split_spaces else r'[,;\u3001\uff0c\uff1b]+'
    text = re.sub(pattern, ',', text)
    return [part.strip() for part in text.split(',') if part.strip()]


def normalize_maker(value):
    text = _cell_str(value).upper()
    text = re.sub(r'\([^)]*\)|\uff08[^\uff09]*\uff09', '', text)
    for suffix in ('\u6709\u9650\u516c\u53f8', '\u80a1\u4efd\u516c\u53f8', '\u516c\u53f8', '\u96c6\u56e2', 'CORPORATION', 'CO.', 'CO', 'INC.', 'INC'):
        if text.endswith(suffix):
            text = text[:-len(suffix)]
    return ''.join(text.split())


def mapped_maker(value):
    text = _cell_str(value)
    if not text:
        return ''
    match = lookup_manufacturer(text)
    return _cell_str(match.get('canonical_name')) if match else text


def identity_key(models, makers):
    model_key = '/'.join(sorted(_cell_str(v).upper().replace(' ', '') for v in models if _cell_str(v)))
    maker_key = '/'.join(sorted(normalize_maker(mapped_maker(v)) for v in makers if _cell_str(v)))
    if not model_key or not maker_key:
        return ''
    return f'\u578b\u53f7:{model_key} | \u5236\u9020\u5546:{maker_key}'


def refdes_key(value):
    refs = split_tokens(value, split_spaces=True)
    return ','.join(sorted(refs)), refs


def pick_col(headers, candidates, normalize_header):
    normalized = {normalize_header(h): h for h in headers if h}
    for cand in candidates:
        found = normalized.get(normalize_header(cand))
        if found:
            return found
    for header in headers:
        norm = normalize_header(header)
        if any(normalize_header(cand) in norm for cand in candidates):
            return header
    return ''


def _preview_status(key, model_values, maker_values):
    issues = []
    if not key:
        issues.append('\u5339\u914d\u952e\u4e3a\u7a7a')
    if not model_values:
        issues.append('\u578b\u53f7\u4e3a\u7a7a')
    if not maker_values:
        issues.append('\u5236\u9020\u5546\u4e3a\u7a7a')
    return ('\u5f02\u5e38' if issues else '\u6709\u6548'), '\uff1b'.join(issues)


def normalize_customer(path, sheet_name, header_row, mapping, match_mode, helpers):
    wb = _open_workbook(path, data_only=True)
    try:
        sheet_name = helpers['pick_sheet'](wb, sheet_name)
        ws = wb[sheet_name]
        headers = helpers['headers_fn'](ws, header_row)
        model_col = _cell_str(mapping.get('model'))
        maker_col = _cell_str(mapping.get('manufacturer'))
        refdes_col = _cell_str(mapping.get('refdes'))
        qty_col = _cell_str(mapping.get('quantity'))
        customer_part_col = _cell_str(mapping.get('customer_part'))
        missing = []
        for label, col in [('\u89c4\u683c\u578b\u53f7', model_col), ('\u5236\u9020\u5546', maker_col)]:
            if not col or col not in headers:
                missing.append(label)
        if match_mode == 'refdes' and (not refdes_col or refdes_col not in headers):
            missing.append('\u4f4d\u53f7')
        if missing:
            raise ValueError('\u8bf7\u5b8c\u6210\u5ba2\u6237 BOM \u5217\u6620\u5c04\uff1a' + missing[0])

        idx = {name: headers.index(col) + 1 for name, col in [
            ('model', model_col),
            ('manufacturer', maker_col),
            ('refdes', refdes_col),
            ('quantity', qty_col),
            ('customer_part', customer_part_col),
        ] if col and col in headers}
        rows = []
        for ri in range(header_row + 1, ws.max_row + 1):
            raw = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, ws.max_column + 1)]
            if not any(raw):
                continue
            models = split_tokens(ws.cell(row=ri, column=idx['model']).value, split_spaces=False)
            makers = split_tokens(ws.cell(row=ri, column=idx['manufacturer']).value, split_spaces=False)
            mapped = [mapped_maker(v) for v in makers]
            ref_text = _cell_str(ws.cell(row=ri, column=idx['refdes']).value) if 'refdes' in idx else ''
            if match_mode == 'refdes':
                key, refs = refdes_key(ref_text)
            else:
                refs = []
                key = identity_key(models, makers)
            status, issue = _preview_status(key, models, makers)
            rows.append({
                'row': ri,
                'match_key': key,
                'refdes': ','.join(sorted(refs)) if refs else ref_text,
                'model': ';'.join(models),
                'manufacturer_raw': ';'.join(makers),
                'manufacturer_mapped': ';'.join(mapped),
                'quantity': _cell_str(ws.cell(row=ri, column=idx['quantity']).value) if 'quantity' in idx else '',
                'customer_part': _cell_str(ws.cell(row=ri, column=idx['customer_part']).value) if 'customer_part' in idx else '',
                'status': status,
                'issue': issue,
            })
        return rows
    finally:
        wb.close()


def normalize_hq(path, sheet_name, match_mode, helpers):
    wb, sheet_name, ws, fmt = helpers['open_hq_info'](path, sheet_name)
    try:
        headers = fmt['headers']
        header_row = fmt['header_row']
        normalize_header = helpers['normalize_header']
        ref_col = pick_col(headers, ['\u4f4d\u53f7', 'refdes', 'reference'], normalize_header)
        model_col = pick_col(headers, ['\u578b\u53f7', '\u89c4\u683c\u578b\u53f7'], normalize_header)
        maker_col = pick_col(headers, ['\u751f\u4ea7\u5382\u5bb6', '\u5236\u9020\u5546', '\u5382\u5546'], normalize_header)
        qty_col = pick_col(headers, ['\u5355\u8017', '\u6570\u91cf', 'qty'], normalize_header)
        part_col = pick_col(headers, ['\u6599\u53f7', 'partnumber', 'part_number'], normalize_header)
        alt_col = pick_col(headers, ['\u66ff\u4ee3\u5173\u7cfb'], normalize_header)
        required = [('\u578b\u53f7', model_col), ('\u751f\u4ea7\u5382\u5bb6', maker_col)]
        if match_mode == 'refdes':
            required.insert(0, ('\u4f4d\u53f7', ref_col))
        missing = [label for label, col in required if not col]
        if missing:
            raise ValueError('HQ BOM \u7f3a\u5c11\u6807\u51c6\u5217\uff1a' + missing[0])
        idx = {name: headers.index(col) + 1 for name, col in [
            ('refdes', ref_col),
            ('model', model_col),
            ('manufacturer', maker_col),
            ('quantity', qty_col),
            ('part_no', part_col),
            ('alternate', alt_col),
        ] if col and col in headers}
        rows = []
        for ri in range(header_row + 1, ws.max_row + 1):
            raw = [_cell_str(ws.cell(row=ri, column=ci).value) for ci in range(1, len(headers) + 1)]
            if fmt['kind'] == 'plm_full' and helpers['is_plm_history_header'](raw):
                break
            if not any(raw):
                continue
            models = split_tokens(ws.cell(row=ri, column=idx['model']).value, split_spaces=False)
            makers = split_tokens(ws.cell(row=ri, column=idx['manufacturer']).value, split_spaces=False)
            mapped = [mapped_maker(v) for v in makers]
            ref_text = _cell_str(ws.cell(row=ri, column=idx['refdes']).value) if 'refdes' in idx else ''
            if match_mode == 'refdes':
                key, refs = refdes_key(ref_text)
            else:
                refs = []
                key = identity_key(models, makers)
            status, issue = _preview_status(key, models, makers)
            rows.append({
                'row': ri,
                'match_key': key,
                'refdes': ','.join(sorted(refs)) if refs else ref_text,
                'model': ';'.join(models),
                'manufacturer_raw': ';'.join(makers),
                'manufacturer_mapped': ';'.join(mapped),
                'quantity': _cell_str(ws.cell(row=ri, column=idx['quantity']).value) if 'quantity' in idx else '',
                'part_no': _cell_str(ws.cell(row=ri, column=idx['part_no']).value) if 'part_no' in idx else '',
                'alternate': _cell_str(ws.cell(row=ri, column=idx['alternate']).value) if 'alternate' in idx else '',
                'status': status,
                'issue': issue,
            })
        return rows
    finally:
        wb.close()


def preview(left_path, right_path, left_sheet, right_sheet, left_header_row, mapping, match_mode, helpers, limit=200):
    if match_mode not in ('refdes', 'identity'):
        raise ValueError('\u8bf7\u9009\u62e9\u5339\u914d\u6a21\u5f0f')
    customer_rows = normalize_customer(left_path, left_sheet, left_header_row, mapping, match_mode, helpers)
    hq_rows = normalize_hq(right_path, right_sheet, match_mode, helpers)
    return {
        'match_mode': match_mode,
        'customer_total': len(customer_rows),
        'hq_total': len(hq_rows),
        'customer_invalid': sum(1 for r in customer_rows if r['status'] != '\u6709\u6548'),
        'hq_invalid': sum(1 for r in hq_rows if r['status'] != '\u6709\u6548'),
        'customer_preview': customer_rows[:limit],
        'hq_preview': hq_rows[:limit],
        'preview_limit': limit,
    }
