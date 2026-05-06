# -*- coding: utf-8 -*-
"""Feishu BOM online sync, mapping suggestion, and local cache API routes."""

from __future__ import annotations

import json
import os
import uuid
from typing import Optional


def _truthy(value) -> bool:
    return str(value or '').strip().lower() in {'1', 'true', 'yes', 'on'}


def _extract_json_object_from_text(text: str) -> dict:
    raw = str(text or '').strip()
    if not raw:
        return {}
    if raw.startswith('```'):
        raw = raw.strip('`').strip()
        if raw.lower().startswith('json'):
            raw = raw[4:].strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find('{')
        end = raw.rfind('}')
        if start >= 0 and end > start:
            return json.loads(raw[start:end + 1])
        raise


def register_feishu_routes(
    app,
    *,
    request,
    jsonify,
    FeishuBomError,
    build_feishu_bom_status,
    build_feishu_database_overview,
    get_feishu_cache_rows,
    create_feishu_cache_row,
    update_feishu_cache_row,
    delete_feishu_cache_library,
    delete_feishu_cache_row,
    fetch_feishu_sheet_list,
    preview_feishu_sheet,
    get_saved_feishu_field_order,
    suggest_feishu_mapping_from_preview,
    build_feishu_mapping_from_headers,
    build_aster_status,
    ask_aster_model,
    AsterHarnessModelProvider,
    sync_feishu_library,
) -> None:
    """Register Feishu BOM routes."""

    @app.get('/api/feishu-bom/status')
    def feishu_bom_status():
        load_runtime = str(request.args.get('runtime', '1')).strip().lower() not in {'0', 'false', 'no'}
        return jsonify(build_feishu_bom_status(load_runtime=load_runtime))

    @app.get('/api/feishu-bom/database')
    def feishu_bom_database():
        return jsonify(build_feishu_database_overview())

    @app.get('/api/feishu-bom/database/rows')
    def feishu_bom_database_rows():
        try:
            raw_limit = str(request.args.get('limit') or '100').strip().lower()
            limit = 5000 if raw_limit in {'all', 'full'} else int(raw_limit or 100)
            result = get_feishu_cache_rows(
                lib_id=str(request.args.get('lib_id') or '').strip(),
                sheet_name=str(request.args.get('sheet_name') or '').strip(),
                query=str(request.args.get('query') or '').strip(),
                limit=limit,
                offset=int(request.args.get('offset') or 0),
            )
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书缓存行读取失败：{exc}'}), 500
        return jsonify(result), 200 if result.get('ok') else 400

    @app.post('/api/feishu-bom/database/rows')
    def feishu_bom_database_create_row():
        try:
            result = create_feishu_cache_row(request.get_json(silent=True) or {})
        except FeishuBomError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书缓存行新增失败：{exc}'}), 500
        return jsonify(result), 201 if result.get('ok') else 400

    @app.patch('/api/feishu-bom/database/rows/<int:row_id>')
    def feishu_bom_database_update_row(row_id: int):
        try:
            result = update_feishu_cache_row(row_id, request.get_json(silent=True) or {})
        except FeishuBomError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书缓存行更新失败：{exc}'}), 500
        return jsonify(result), 200 if result.get('ok') else 404

    @app.delete('/api/feishu-bom/database/libraries/<lib_id>')
    def feishu_bom_database_delete_library(lib_id: str):
        try:
            result = delete_feishu_cache_library(lib_id)
        except FeishuBomError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书缓存库删除失败：{exc}'}), 500
        return jsonify(result)

    @app.delete('/api/feishu-bom/database/rows/<int:row_id>')
    def feishu_bom_database_delete_row(row_id: int):
        try:
            result = delete_feishu_cache_row(row_id)
        except FeishuBomError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书缓存行删除失败：{exc}'}), 500
        return jsonify(result), 200 if result.get('ok') else 404

    @app.post('/api/feishu-bom/sheets')
    def feishu_bom_sheets():
        data = request.get_json(silent=True) or request.form.to_dict()
        token_or_url = str(
            data.get('spreadsheet_token_or_url')
            or data.get('spreadsheetToken')
            or data.get('token')
            or ''
        ).strip()
        if not token_or_url:
            return jsonify({'ok': False, 'error': '请提供 spreadsheet_token_or_url。'}), 400
        try:
            result = fetch_feishu_sheet_list(
                spreadsheet_token_or_url=token_or_url,
                base_url=str(data.get('base_url') or '').strip(),
                origin=str(data.get('origin') or '').strip(),
                user_id=str(data.get('user_id') or data.get('userId') or '').strip(),
            )
        except FeishuBomError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书 Sheet 列表获取失败：{exc}'}), 500
        return jsonify(result)

    @app.post('/api/feishu-bom/preview-sheet')
    def feishu_bom_preview_sheet():
        data = request.get_json(silent=True) or request.form.to_dict()
        token_or_url = str(
            data.get('spreadsheet_token_or_url')
            or data.get('spreadsheetToken')
            or data.get('token')
            or ''
        ).strip()
        if not token_or_url:
            return jsonify({'ok': False, 'error': '请提供 spreadsheet_token_or_url。'}), 400
        try:
            result = preview_feishu_sheet(
                spreadsheet_token_or_url=token_or_url,
                sheet_id=str(data.get('sheet_id') or data.get('sheetId') or '').strip(),
                base_url=str(data.get('base_url') or '').strip(),
                origin=str(data.get('origin') or '').strip(),
                user_id=str(data.get('user_id') or data.get('userId') or '').strip(),
                row_count=int(data.get('row_count') or 50),
                column_range=str(data.get('column_range') or data.get('col_range') or 'A:Z').strip(),
                header_row=int(data.get('header_row') or 1),
            )
        except FeishuBomError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书 Sheet 预览失败：{exc}'}), 500
        return jsonify(result)

    @app.post('/api/feishu-bom/suggest-mapping')
    def feishu_bom_suggest_mapping():
        data = request.get_json(silent=True) or request.form.to_dict()
        rows = data.get('rows') or data.get('values') or []
        if isinstance(rows, str):
            try:
                rows = json.loads(rows)
            except json.JSONDecodeError:
                return jsonify({'ok': False, 'error': 'rows 必须是二维 JSON 数组。'}), 400
        if not isinstance(rows, list):
            return jsonify({'ok': False, 'error': 'rows 必须是数组。'}), 400
        sheet_title = str(data.get('sheet_title') or data.get('title') or '').strip()
        saved_order = get_saved_feishu_field_order().get('optional_field_order', [])
        header_row = None
        if data.get('header_row') not in {None, ''}:
            try:
                header_row = int(data.get('header_row') or 1)
            except (TypeError, ValueError):
                return jsonify({'ok': False, 'error': 'header_row 必须是数字。'}), 400
        suggestion = suggest_feishu_mapping_from_preview(
            rows,
            sheet_title=sheet_title,
            saved_optional_order=saved_order,
            header_row=header_row,
        )
        result = {
            'ok': True,
            'mode': 'local-heuristic',
            'suggestion': suggestion,
            'saved_field_order': saved_order,
            'agent': {
                'used': False,
                'status': build_aster_status(),
                'role': 'header-title-and-optional-field-suggestion',
                'error': '',
            },
        }
        if not _truthy(data.get('use_agent')):
            return jsonify(result)

        status_payload = result['agent']['status']
        if status_payload.get('mode') in {'', 'mock', 'off'}:
            result['agent']['error'] = 'Aster 当前不是 live 模式，已返回本地启发式建议。'
            return jsonify(result)

        prompt = (
            '你是硬件 BOM 优选库表格表头识别助手。请只输出一个 JSON 对象，不要 Markdown、不要解释段落。'
            '本次请求只代表一个 Sheet；必须仅基于本次 preview_rows 独立判断，禁止引用历史会话、其他 Sheet、前一次回答或外部常识补全不存在的列。'
            '你的任务不是同步数据，也不是匹配物料，只识别：1) 哪一行是表头；2) 该行原始表头 title 列表；3) 值得保留的扩展字段 title。'
            '必须保留表格中的原始 title 文本，不要翻译、改名或创造列名。'
            '标准字段识别目标包括：HQ料号/物料编码/HQ编码等同义列，特征通常是 HQ 开头的一串数字或编码；规格型号/Part Number/厂家型号/制造商型号；PI；选型顺序/优选顺序/priority/rank。'
            '其中选型顺序可能不存在，尤其芯片表可以为空；不存在时不要硬猜。'
            '扩展字段 optional_titles 只输出当前表头中确实存在的 title，可参考 saved_field_order 的历史顺序优先排列，再补充当前物料族明显有价值的字段，例如封装、耐压、容量、功率、精度、材质、品牌。'
            '不要把空列、序号列、备注性大段说明列、图片/链接列误判为标准字段。'
            'JSON 字段必须包含 header_row、headers、optional_titles、confidence、notes。'
            'header_row 从 1 开始计数；headers 是该表头行的完整 title 字符串数组；optional_titles 是扩展字段 title 数组；confidence 只能是 high/medium/low。'
            '不能确定时 confidence=low，并在 notes 中说明需要人工确认的原因。'
        )
        sheet_agent_request_id = uuid.uuid4().hex[:12]
        inputs = {
            'sheet_agent_request_id': sheet_agent_request_id,
            'sheet_title': sheet_title,
            'preview_rows': rows[:16],
            'saved_field_order': saved_order,
            'required_targets': ['HQ料号', '规格型号/Part Number', 'PI', '选型顺序'],
            'output_contract': {
                'header_row': '1-based integer',
                'headers': 'exact titles copied from the detected header row',
                'optional_titles': 'exact existing titles only',
                'confidence': 'high|medium|low',
                'notes': 'array of short Chinese notes',
            },
            'isolation': {
                'conversation_id': '',
                'with_context': False,
                'scope': 'current_sheet_only',
            },
        }
        try:
            isolated_env = dict(os.environ)
            isolated_env['PSTX_ASTER_CONVERSATION_ID'] = ''
            isolated_env['PSTX_ASTER_AUTO_GENERATE_NAME'] = 'false'

            def isolated_ask_model(model_prompt: str, *, inputs: Optional[dict] = None) -> dict:
                return ask_aster_model(model_prompt, inputs=inputs or {}, environ=isolated_env)

            provider = AsterHarnessModelProvider(ask_model=isolated_ask_model)
            response = provider.generate(prompt, inputs=inputs)
            parsed = _extract_json_object_from_text(response.answer)
            if isinstance(parsed, dict):
                header_row = int(parsed.get('header_row') or suggestion.get('header_row') or 1)
                headers = list(parsed.get('headers') or suggestion.get('headers') or [])
                mapped = build_feishu_mapping_from_headers(
                    headers,
                    header_row=header_row,
                    sheet_title=sheet_title,
                    provider=response.provider,
                    saved_optional_order=saved_order,
                    optional_titles=list(parsed.get('optional_titles') or []),
                    notes=[
                        'Agent 识别表头行、表头 title 和扩展字段候选；标准字段由本地规则按这些表头生成草稿，仍需人工确认。',
                        *list(parsed.get('notes') or []),
                    ],
                )
                mapped['header_detection'] = {
                    'provider': response.provider,
                    'header_row': header_row,
                    'headers': headers,
                    'optional_titles': parsed.get('optional_titles', []),
                    'confidence': parsed.get('confidence', ''),
                    'notes': parsed.get('notes', []),
                }
                result['mode'] = 'agent-assisted'
                result['suggestion'] = mapped
                result['agent'].update({
                    'used': True,
                    'provider': response.provider,
                    'mode': response.mode,
                    'metadata': {
                        **response.metadata,
                        'isolated_conversation': True,
                        'sheet_agent_request_id': sheet_agent_request_id,
                    },
                })
        except Exception as exc:
            result['agent']['error'] = f'Agent 字段建议失败，已回退本地启发式：{exc}'
        return jsonify(result)

    @app.post('/api/feishu-bom/sync')
    def feishu_bom_sync():
        data = request.get_json(silent=True) or request.form.to_dict()
        token_or_url = str(
            data.get('spreadsheet_token_or_url')
            or data.get('spreadsheetToken')
            or data.get('token')
            or ''
        ).strip()
        sheets = data.get('sheets') or []
        if isinstance(sheets, str):
            try:
                sheets = json.loads(sheets)
            except json.JSONDecodeError:
                return jsonify({'ok': False, 'error': 'sheets 必须是 JSON 数组。'}), 400
        if not isinstance(sheets, list):
            return jsonify({'ok': False, 'error': 'sheets 必须是数组。'}), 400
        try:
            result = sync_feishu_library(
                library_name=str(data.get('library_name') or data.get('name') or '').strip(),
                spreadsheet_token_or_url=token_or_url,
                library_id=str(data.get('library_id') or '').strip(),
                sheets=sheets,
                base_url=str(data.get('base_url') or '').strip(),
                origin=str(data.get('origin') or '').strip(),
                user_id=str(data.get('user_id') or data.get('userId') or '').strip(),
            )
        except FeishuBomError as exc:
            return jsonify({'ok': False, 'error': str(exc)}), 400
        except Exception as exc:
            return jsonify({'ok': False, 'error': f'飞书 BOM 同步失败：{exc}'}), 500
        return jsonify(result)
