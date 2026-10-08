"""Load an API definition from CSV or Excel."""

from __future__ import annotations

import csv
import io
import os
from typing import Dict, List, Optional, Sequence, Tuple

from jmeter_gen.model import (
    DEFAULT_PLAN_NAME,
    ApiRequest,
    InputError,
    Plan,
    Settings,
    Sheet,
    normalize_token,
    parse_bool,
    parse_int_range,
    require_xml_text,
    validate_header_name,
    validate_header_value,
    validate_param_name,
    validate_relative_jtl,
)
from jmeter_gen.model import ON_ERROR_VALUES
from jmeter_gen.xlsxio import read_xlsx, write_xlsx

STRUCTURED_ROLES = [
    ("record_type", ("record_type", "レコード種別", "種別")),
    ("api_name", ("api_name", "apiname", "api名", "名前", "名称", "name")),
    ("enabled", ("enabled", "enable", "有効")),
    ("method", ("method", "http_method", "メソッド", "httpメソッド")),
    ("url", ("url", "uri", "endpoint", "アドレス")),
    ("body", ("body", "request_body", "ボディ", "リクエストボディ")),
    ("content_type", ("content_type", "contenttype", "コンテントタイプ")),
    ("body_mode", ("body_mode", "bodymode", "ボディ形式")),
    ("timeout_ms", ("timeout_ms", "timeout", "タイムアウトms", "タイムアウト")),
    ("expected_status", ("expected_status", "期待ステータス", "期待するステータス")),
    ("key", ("key", "キー", "パラメータ名", "ヘッダ名")),
    ("value", ("value", "値")),
    ("comment", ("comment", "comments", "コメント", "説明", "備考")),
]

API_ROLES = [
    ("enabled", ("enabled", "enable", "有効")),
    ("name", ("name", "api_name", "apiname", "名前", "名称", "api名")),
    ("method", ("method", "http_method", "メソッド", "httpメソッド")),
    ("url", ("url", "uri", "endpoint", "アドレス")),
    ("body", ("body", "request_body", "ボディ", "リクエストボディ")),
    ("content_type", ("content_type", "contenttype", "コンテントタイプ")),
    ("body_mode", ("body_mode", "bodymode", "ボディ形式")),
    ("timeout_ms", ("timeout_ms", "timeout", "タイムアウトms", "タイムアウト")),
    ("expected_status", ("expected_status", "期待ステータス", "期待するステータス")),
    ("comment", ("comment", "comments", "コメント", "説明", "備考")),
]

CHILD_ROLES = [
    ("api_name", ("api_name", "apiname", "api名", "対象api", "対象")),
    ("key", ("key", "param_name", "name", "名前", "キー", "パラメータ名", "ヘッダ名")),
    ("value", ("value", "param_value", "値")),
    ("enabled", ("enabled", "enable", "有効")),
]

SETTING_ROLES = [
    ("key", ("key", "name", "名前", "キー", "項目")),
    ("value", ("value", "値", "内容")),
    ("comment", ("comment", "コメント", "説明", "備考")),
]

FLAT_ROLES = [
    ("name", ("name", "api_name", "名前", "名称", "api名")),
    ("method", ("method", "http_method", "メソッド", "httpメソッド")),
    ("url", ("url", "uri", "endpoint", "アドレス")),
    ("query", ("query", "query_params", "クエリ", "クエリパラメータ")),
    ("headers", ("headers", "header", "ヘッダ", "ヘッダー", "ヘッダ情報")),
    ("body", ("body", "request_body", "ボディ", "リクエストボディ")),
    ("form", ("form", "form_params", "フォーム", "postパラメータ")),
    ("content_type", ("content_type", "contenttype", "コンテントタイプ")),
    ("body_mode", ("body_mode", "bodymode", "ボディ形式")),
    ("timeout_ms", ("timeout_ms", "timeout", "タイムアウトms", "タイムアウト")),
    ("expected_status", ("expected_status", "期待ステータス", "期待するステータス")),
    ("comment", ("comment", "comments", "コメント", "説明", "備考")),
    ("enabled", ("enabled", "enable", "有効")),
]

SHEET_KIND = {
    "api": "api",
    "apis": "api",
    "リクエスト": "api",
    "query": "query",
    "queryparams": "query",
    "クエリ": "query",
    "クエリパラメータ": "query",
    "header": "header",
    "headers": "header",
    "ヘッダ": "header",
    "ヘッダー": "header",
    "ヘッダ情報": "header",
    "body": "form",
    "bodyparams": "form",
    "form": "form",
    "forms": "form",
    "フォーム": "form",
    "postパラメータ": "form",
    "ボディパラメータ": "form",
    "setting": "settings",
    "settings": "settings",
    "設定": "settings",
}

IGNORED_SHEETS = {"readme", "説明", "読み方", "使い方", "表紙", "manual"}

RECORD_KIND = {
    "api": "API",
    "query": "QUERY",
    "header": "HEADER",
    "body": "BODY",
    "form": "BODY",
    "setting": "SETTING",
    "クエリ": "QUERY",
    "ヘッダ": "HEADER",
    "ヘッダー": "HEADER",
    "フォーム": "BODY",
    "ボディ": "BODY",
    "設定": "SETTING",
}

SETTING_FIELDS = {
    "threads": "threads",
    "thread": "threads",
    "スレッド数": "threads",
    "スレッド": "threads",
    "loops": "loops",
    "loop": "loops",
    "ループ回数": "loops",
    "ループ": "loops",
    "ramp_up_sec": "ramp_up_sec",
    "ramp_up": "ramp_up_sec",
    "rampup": "ramp_up_sec",
    "ランプアップ秒": "ramp_up_sec",
    "ランプアップ": "ramp_up_sec",
    "follow_redirects": "follow_redirects",
    "followredirects": "follow_redirects",
    "リダイレクト追従": "follow_redirects",
    "plan_name": "plan_name",
    "planname": "plan_name",
    "プラン名": "plan_name",
    "on_error": "on_error",
    "onerror": "on_error",
    "エラー時": "on_error",
    "jtl_file": "jtl_file",
    "jtlファイル": "jtl_file",
}


def _sheet_kind(name: str) -> str:
    token = normalize_token(name)
    if token in IGNORED_SHEETS:
        return "ignore"
    return SHEET_KIND.get(token, "unknown")


def _map_headers(
    headers: Sequence[str],
    roles: Sequence[Tuple[str, Sequence[str]]],
) -> Tuple[Dict[str, int], List[str]]:
    normalized = [normalize_token(header) for header in headers]
    used = set()
    found: Dict[str, int] = {}
    for role, aliases in roles:
        chosen = None
        for alias in aliases:
            token = normalize_token(alias)
            for index, header in enumerate(normalized):
                if index in used or header != token:
                    continue
                chosen = index
                break
            if chosen is not None:
                break
        if chosen is not None:
            used.add(chosen)
            found[role] = chosen
    unused = [
        str(headers[index])
        for index in range(len(headers))
        if index not in used and str(headers[index]).strip() != ""
    ]
    return found, unused


def _row_values(mapping: Dict[str, int], row: Sequence[str]) -> Dict[str, str]:
    values: Dict[str, str] = {}
    for role, index in mapping.items():
        if index < len(row) and row[index] is not None:
            values[role] = str(row[index])
        else:
            values[role] = ""
    return values


def _is_blank(row: Sequence[str]) -> bool:
    return all(str(cell).strip() == "" for cell in row)


def _require_columns(mapping: Dict[str, int], required: Sequence[str], label: str) -> None:
    missing = [name for name in required if name not in mapping]
    if missing:
        raise InputError(f"{label} に必須列がありません: {', '.join(missing)}")


def _record_kind(value: str, label: str) -> str:
    token = normalize_token(value)
    if token == "":
        raise InputError(f"{label}: record_type が空です。")
    kind = RECORD_KIND.get(token)
    if kind is None:
        raise InputError(f"{label}: record_type が不正です: {value}")
    return kind


def _blank_record() -> Dict[str, str]:
    return {role: "" for role, _ in STRUCTURED_ROLES if role != "record_type"}


def _iter_table(sheet: Sheet) -> List[Tuple[str, List[str]]]:
    numbered: List[Tuple[str, List[str]]] = []
    numbers = sheet.row_numbers or list(range(1, len(sheet.rows) + 1))
    for index, row in enumerate(sheet.rows):
        number = numbers[index] if index < len(numbers) else index + 1
        numbered.append((f"{sheet.name} {number}行目", list(row)))
    return numbered


def _header_and_data(
    table: Sequence[Tuple[str, List[str]]],
    label: str,
) -> Tuple[str, List[str], List[Tuple[str, List[str]]]]:
    for position, (row_label, row) in enumerate(table):
        if _is_blank(row):
            continue
        return row_label, row, list(table[position + 1 :])
    raise InputError(f"{label} にヘッダ行がありません。")


def parse_csv_text(text: str, source_path: str) -> Plan:
    warnings: List[str] = []
    reader = csv.reader(io.StringIO(text, newline=""))
    rows = [(f"CSV {number}行目", list(row)) for number, row in enumerate(reader, start=1)]
    if not rows or all(_is_blank(row) for _, row in rows):
        raise InputError("CSVが空です。")
    _, header, data = _header_and_data(rows, "CSV")
    mapping, unused = _map_headers(header, STRUCTURED_ROLES)
    for name in unused:
        warnings.append(f"CSV: 未使用の列を無視しました: {name}")
    if "record_type" in mapping:
        _require_columns(mapping, ("record_type",), "CSV")
        records = []
        for row_label, row in data:
            if _is_blank(row):
                continue
            values = _row_values(mapping, row)
            values["kind"] = _record_kind(values.get("record_type", ""), row_label)
            values["row_label"] = row_label
            records.append(values)
        return _plan_from_records(records, warnings, source_path)
    flat_map, flat_unused = _map_headers(header, FLAT_ROLES)
    if "method" in flat_map and "url" in flat_map and "name" in flat_map:
        for name in flat_unused:
            warnings.append(f"CSV: 未使用の列を無視しました: {name}")
        return _plan_from_flat(data, flat_map, warnings, source_path)
    raise InputError(
        "CSVのヘッダを認識できません。record_type 列か、name,method,url 列が必要です。"
    )


def _plan_from_flat(
    data: Sequence[Tuple[str, List[str]]],
    mapping: Dict[str, int],
    warnings: List[str],
    source_path: str,
) -> Plan:
    records = []
    for row_label, row in data:
        if _is_blank(row):
            continue
        values = _row_values(mapping, row)
        api_values = _blank_record()
        api_values.update(
            {
                "api_name": values.get("name", ""),
                "enabled": values.get("enabled", ""),
                "method": values.get("method", ""),
                "url": values.get("url", ""),
                "body": values.get("body", ""),
                "content_type": values.get("content_type", ""),
                "body_mode": values.get("body_mode", ""),
                "timeout_ms": values.get("timeout_ms", ""),
                "expected_status": values.get("expected_status", ""),
                "comment": values.get("comment", ""),
                "kind": "API",
                "row_label": row_label,
            }
        )
        records.append(api_values)
        api_name = values.get("name", "").strip()
        for key, value in _parse_query_blob(values.get("query", ""), f"{row_label} のクエリ"):
            child = _blank_record()
            child.update(
                {
                    "kind": "QUERY",
                    "api_name": api_name,
                    "key": key,
                    "value": value,
                    "enabled": "true",
                    "row_label": row_label,
                }
            )
            records.append(child)
        for key, value in _parse_query_blob(values.get("form", ""), f"{row_label} のフォーム"):
            child = _blank_record()
            child.update(
                {
                    "kind": "BODY",
                    "api_name": api_name,
                    "key": key,
                    "value": value,
                    "enabled": "true",
                    "row_label": row_label,
                }
            )
            records.append(child)
        for key, value in _parse_headers_blob(values.get("headers", "")):
            child = _blank_record()
            child.update(
                {
                    "kind": "HEADER",
                    "api_name": api_name,
                    "key": key,
                    "value": value,
                    "enabled": "true",
                    "row_label": row_label,
                }
            )
            records.append(child)
    return _plan_from_records(records, warnings, source_path)


def _parse_query_blob(text: str, label: str) -> List[Tuple[str, str]]:
    raw = (text or "").strip()
    if raw == "":
        return []
    try:
        pairs = parse_qsl_strict(raw)
    except ValueError as exc:
        raise InputError(
            f"{label} は key=value&key2=value2 の形式で指定してください。"
        ) from exc
    result: List[Tuple[str, str]] = []
    for key, value in pairs:
        result.append((validate_param_name(key, label), require_xml_text(value, label)))
    return result


def parse_qsl_strict(raw: str) -> List[Tuple[str, str]]:
    from urllib.parse import parse_qsl

    pairs = parse_qsl(
        raw,
        keep_blank_values=True,
        strict_parsing=True,
        encoding="utf-8",
        errors="strict",
    )
    return [(key, value) for key, value in pairs]


def _parse_headers_blob(text: str) -> List[Tuple[str, str]]:
    raw = (text or "").strip()
    if raw == "":
        return []
    result: List[Tuple[str, str]] = []
    for part in split_header_blob(raw):
        item = part.strip()
        if item == "":
            continue
        if ":" not in item:
            raise InputError(
                "ヘッダは '名前: 値' です。複数指定するときは | か改行で区切ってください: "
                + item
            )
        name, value = item.split(":", 1)
        result.append((validate_header_name(name), validate_header_value(value)))
    return result


def split_header_blob(raw: str) -> List[str]:
    import re

    return re.split(r"[|\r\n]+", raw)


def _plan_from_records(
    records: Sequence[Dict[str, str]],
    warnings: List[str],
    source_path: str,
) -> Plan:
    apis: List[ApiRequest] = []
    index: Dict[str, ApiRequest] = {}
    for record in records:
        if record.get("kind") != "API":
            continue
        label = record.get("row_label", "API")
        name = record.get("api_name", "").strip()
        if name == "":
            raise InputError(f"{label}: API名が空です。")
        if len(name) > 200:
            raise InputError(f"{label}: API名が長すぎます。")
        require_xml_text(name, f"{label} のAPI名")
        if name in index:
            raise InputError(f"API名が重複しています: {name}")
        method = record.get("method", "").strip().upper()
        if method not in ("GET", "POST"):
            raise InputError(f"{label}: メソッドは GET または POST のみです: {record.get('method', '')}")
        url = record.get("url", "").strip()
        if url == "":
            raise InputError(f"{label}: URLが空です。")
        api = ApiRequest(
            name=name,
            method=method,
            url=url,
            body=record.get("body", ""),
            content_type=record.get("content_type", "").strip(),
            body_mode=record.get("body_mode", "").strip() or "auto",
            timeout_ms=record.get("timeout_ms", "").strip(),
            expected_status=record.get("expected_status", "").strip(),
            comment=record.get("comment", "").strip(),
            enabled=parse_bool(record.get("enabled", ""), default=True, label=f"{label} の有効"),
            row_label=label,
        )
        require_xml_text(api.body, f"{name} のボディ")
        require_xml_text(api.comment, f"{name} のコメント")
        apis.append(api)
        index[name] = api
    if not apis:
        raise InputError("APIが1件もありません。")
    settings = Settings()
    seen_settings = set()
    for record in records:
        kind = record.get("kind")
        label = record.get("row_label", kind or "行")
        if kind == "API":
            continue
        if kind == "SETTING":
            _apply_setting(settings, seen_settings, record, label)
            continue
        api_name = record.get("api_name", "").strip()
        if api_name not in index:
            raise InputError(f"{label}: 対象APIが見つかりません: {api_name}")
        if not parse_bool(record.get("enabled", ""), default=True, label=f"{label} の有効"):
            continue
        api = index[api_name]
        key = record.get("key", "")
        value = record.get("value", "")
        if kind == "QUERY":
            api.query.append((validate_param_name(key, f"{label} のクエリ名"), require_xml_text(value, f"{label} のクエリ値")))
        elif kind == "HEADER":
            api.headers.append((validate_header_name(key), validate_header_value(value)))
        elif kind == "BODY":
            api.form_params.append(
                (validate_param_name(key, f"{label} のフォーム名"), require_xml_text(value, f"{label} のフォーム値"))
            )
        else:
            raise InputError(f"{label}: 未知のレコードです。")
    return Plan(apis=apis, settings=settings, warnings=warnings, source_path=source_path)


def _apply_setting(
    settings: Settings,
    seen: set,
    record: Dict[str, str],
    label: str,
) -> None:
    raw_key = record.get("key", "").strip()
    if raw_key == "":
        raise InputError(f"{label}: 設定名が空です。")
    field_name = SETTING_FIELDS.get(normalize_token(raw_key))
    if field_name is None:
        raise InputError(f"{label}: 未知の設定です: {raw_key}")
    if field_name in seen:
        raise InputError(f"設定が重複しています: {raw_key}")
    seen.add(field_name)
    value = record.get("value", "")
    if field_name == "threads":
        settings.threads = parse_int_range(value, "threads", 1, 1000000)
    elif field_name == "loops":
        settings.loops = parse_int_range(value, "loops", 1, 1000000)
    elif field_name == "ramp_up_sec":
        settings.ramp_up_sec = parse_int_range(value, "ramp_up_sec", 0, 86400)
    elif field_name == "follow_redirects":
        settings.follow_redirects = parse_bool(value, default=True, label="follow_redirects")
    elif field_name == "plan_name":
        text = str(value).strip()
        if text == "":
            raise InputError("plan_name が空です。")
        require_xml_text(text, "plan_name")
        if len(text) > 200:
            raise InputError("plan_name が長すぎます。")
        settings.plan_name = text
    elif field_name == "on_error":
        mapped = ON_ERROR_VALUES.get(str(value).strip())
        if mapped is None:
            mapped = ON_ERROR_VALUES.get(normalize_token(value))
        if mapped is None:
            raise InputError(f"on_error の値を解釈できません: {value}")
        settings.on_error = mapped
    elif field_name == "jtl_file":
        settings.jtl_file = validate_relative_jtl(value)
    else:
        raise InputError(f"未知の設定です: {raw_key}")


def parse_sheets(sheets: Sequence[Sheet], source_path: str) -> Plan:
    warnings: List[str] = []
    grouped: Dict[str, Sheet] = {}
    loose: List[Sheet] = []
    for sheet in sheets:
        kind = _sheet_kind(sheet.name)
        if kind == "ignore":
            continue
        if kind == "unknown":
            loose.append(sheet)
            continue
        if kind in grouped:
            raise InputError(f"同じ種類のシートが複数あります: {sheet.name}")
        grouped[kind] = sheet
    if "api" in grouped:
        for sheet in loose:
            warnings.append(f"未使用のシートを無視しました: {sheet.name}")
        return _plan_from_workbook(grouped, warnings, source_path)
    if len(loose) == 1:
        return _plan_from_single_sheet(loose[0], warnings, source_path)
    if len(loose) > 1:
        raise InputError("APIシートがありません。シート名は API / クエリ / ヘッダ / フォーム / 設定 にしてください。")
    raise InputError("APIシートがありません。")


def _plan_from_single_sheet(sheet: Sheet, warnings: List[str], source_path: str) -> Plan:
    table = _iter_table(sheet)
    _, header, data = _header_and_data(table, sheet.name)
    mapping, unused = _map_headers(header, STRUCTURED_ROLES)
    if "record_type" in mapping:
        for name in unused:
            warnings.append(f"{sheet.name}: 未使用の列を無視しました: {name}")
        records = []
        for row_label, row in data:
            if _is_blank(row):
                continue
            values = _row_values(mapping, row)
            values["kind"] = _record_kind(values.get("record_type", ""), row_label)
            values["row_label"] = row_label
            records.append(values)
        return _plan_from_records(records, warnings, source_path)
    flat_map, flat_unused = _map_headers(header, FLAT_ROLES)
    if "method" in flat_map and "url" in flat_map and "name" in flat_map:
        for name in flat_unused:
            warnings.append(f"{sheet.name}: 未使用の列を無視しました: {name}")
        return _plan_from_flat(data, flat_map, warnings, source_path)
    raise InputError(f"{sheet.name} のヘッダを認識できません。")


def _records_from_child_sheet(sheet: Sheet, kind: str, warnings: List[str]) -> List[Dict[str, str]]:
    table = _iter_table(sheet)
    if all(_is_blank(row) for _, row in table):
        return []
    _, header, data = _header_and_data(table, sheet.name)
    mapping, unused = _map_headers(header, CHILD_ROLES)
    for name in unused:
        warnings.append(f"{sheet.name}: 未使用の列を無視しました: {name}")
    _require_columns(mapping, ("api_name", "key", "value"), sheet.name)
    records = []
    for row_label, row in data:
        if _is_blank(row):
            continue
        values = _blank_record()
        values.update(_row_values(mapping, row))
        values["kind"] = kind
        values["row_label"] = row_label
        records.append(values)
    return records


def _plan_from_workbook(
    grouped: Dict[str, Sheet],
    warnings: List[str],
    source_path: str,
) -> Plan:
    api_sheet = grouped["api"]
    table = _iter_table(api_sheet)
    _, header, data = _header_and_data(table, api_sheet.name)
    mapping, unused = _map_headers(header, API_ROLES)
    for name in unused:
        warnings.append(f"{api_sheet.name}: 未使用の列を無視しました: {name}")
    _require_columns(mapping, ("name", "method", "url"), api_sheet.name)
    records: List[Dict[str, str]] = []
    for row_label, row in data:
        if _is_blank(row):
            continue
        values = _blank_record()
        got = _row_values(mapping, row)
        values.update(
            {
                "api_name": got.get("name", ""),
                "enabled": got.get("enabled", ""),
                "method": got.get("method", ""),
                "url": got.get("url", ""),
                "body": got.get("body", ""),
                "content_type": got.get("content_type", ""),
                "body_mode": got.get("body_mode", ""),
                "timeout_ms": got.get("timeout_ms", ""),
                "expected_status": got.get("expected_status", ""),
                "comment": got.get("comment", ""),
                "kind": "API",
                "row_label": row_label,
            }
        )
        records.append(values)
    if "query" in grouped:
        records.extend(_records_from_child_sheet(grouped["query"], "QUERY", warnings))
    if "header" in grouped:
        records.extend(_records_from_child_sheet(grouped["header"], "HEADER", warnings))
    if "form" in grouped:
        records.extend(_records_from_child_sheet(grouped["form"], "BODY", warnings))
    if "settings" in grouped:
        records.extend(_setting_records(grouped["settings"], warnings))
    return _plan_from_records(records, warnings, source_path)


def _setting_records(sheet: Sheet, warnings: List[str]) -> List[Dict[str, str]]:
    table = _iter_table(sheet)
    if all(_is_blank(row) for _, row in table):
        return []
    _, header, data = _header_and_data(table, sheet.name)
    mapping, unused = _map_headers(header, SETTING_ROLES)
    for name in unused:
        warnings.append(f"{sheet.name}: 未使用の列を無視しました: {name}")
    _require_columns(mapping, ("key", "value"), sheet.name)
    records = []
    for row_label, row in data:
        if _is_blank(row):
            continue
        values = _blank_record()
        values.update(_row_values(mapping, row))
        values["kind"] = "SETTING"
        values["row_label"] = row_label
        records.append(values)
    return records


def read_text_auto(path: str) -> str:
    with open(path, "rb") as handle:
        raw = handle.read()
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    decoded: Optional[str] = None
    try:
        decoded = raw.decode("utf-8")
    except UnicodeDecodeError:
        decoded = None
    if decoded is not None and _looks_like_table(decoded):
        return decoded
    for encoding in ("cp932", "shift_jis"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    if decoded is not None:
        return decoded
    raise InputError("CSVの文字コードを判定できません。UTF-8で保存してください。")


def _looks_like_table(text: str) -> bool:
    for line in text.splitlines():
        if line.strip():
            return "," in line
    return False


def load_path(path: str) -> Plan:
    if not os.path.isfile(path):
        raise InputError(f"入力ファイルがありません: {path}")
    suffix = os.path.splitext(path)[1].lower()
    if suffix in (".xlsx", ".xlsm"):
        return parse_sheets(read_xlsx(path), os.path.abspath(path))
    if suffix in (".csv", ".txt"):
        return parse_csv_text(read_text_auto(path), os.path.abspath(path))
    raise InputError("入力は .xlsx または .csv にしてください: " + path)


def template_tables() -> Dict[str, List[List[str]]]:
    apis = [
        ["有効", "名前", "メソッド", "URL", "ボディ", "Content-Type", "ボディ形式", "タイムアウトms", "期待ステータス", "コメント"],
        ["true", "get-item", "GET", "https://example.com/v1/items", "", "", "", "", "200", "サンプルGET。実行前にURLを置き換えてください。"],
        ["true", "create-item", "POST", "https://example.com/v1/items", '{"name":"sample","count":1}', "application/json", "raw", "", "", "サンプルPOST(JSON)。"],
        ["true", "submit-form", "POST", "https://example.com/v1/forms", "", "", "form", "", "", "サンプルPOST(フォーム)。"],
    ]
    queries = [
        ["有効", "API名", "名前", "値"],
        ["true", "get-item", "page", "1"],
        ["true", "get-item", "size", "10"],
    ]
    headers = [
        ["有効", "API名", "名前", "値"],
        ["true", "get-item", "Accept", "application/json"],
        ["true", "get-item", "X-Request-Id", "demo-001"],
        ["true", "create-item", "Accept", "application/json"],
        ["true", "submit-form", "Accept", "application/json"],
    ]
    forms = [
        ["有効", "API名", "名前", "値"],
        ["true", "submit-form", "title", "hello"],
        ["true", "submit-form", "note", "world"],
    ]
    settings = [
        ["キー", "値", "コメント"],
        ["threads", "1", "初期値は1スレッドです。各APIを1回呼びます。"],
        ["loops", "1", "初期値は1ループです。"],
        ["ramp_up_sec", "1", ""],
        ["follow_redirects", "true", ""],
        ["plan_name", DEFAULT_PLAN_NAME, ""],
        ["on_error", "continue", "continue / stopthread / stoptest / stoptestnow"],
    ]
    manual = [
        ["項目", "内容"],
        ["このファイル", "JMeter 5.6.3 のテスト計画を作るための入力テンプレートです。"],
        ["編集するシート", "API、クエリ、ヘッダ、フォーム、設定。この説明シートは読み込まれません。"],
        ["API", "1行が1つのHTTPリクエストです。メソッドは GET または POST です。"],
        ["クエリ", "API名は APIシートの名前と一致させます。GETのクエリ、POSTのURLクエリの両方に使います。"],
        ["ヘッダ", "ヘッダ名と値を1行ずつ書きます。Authorization などの秘密値は生成されるJMXに平文で入ります。"],
        ["フォーム", "application/x-www-form-urlencoded の項目です。JSONなどの生ボディとは同時に使えません。"],
        ["ボディ", "JSONなど、そのまま送りたいRequest Bodyです。フォーム項目がある行では空にします。"],
        ["ボディ形式", "空または auto は自動判定です。raw は生ボディ、form はフォームです。"],
        ["初期負荷", "設定シートの threads=1、loops=1 が初期値です。有効な各APIを1回呼びます。"],
        ["期待ステータス", "空なら応答コードでは成否を判定しません。200 のように書くと一致を検証します。"],
        ["無効化", "有効列を false にすると、そのAPIまたは項目をJMXに入れません。"],
    ]
    return {
        "説明": manual,
        "API": apis,
        "クエリ": queries,
        "ヘッダ": headers,
        "フォーム": forms,
        "設定": settings,
    }


def template_csv_text() -> str:
    tables = template_tables()
    apis = tables["API"][1:]
    queries = tables["クエリ"][1:]
    headers = tables["ヘッダ"][1:]
    forms = tables["フォーム"][1:]
    settings = tables["設定"][1:]
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer, lineterminator="\r\n")
    writer.writerow(
        [
            "record_type",
            "api_name",
            "enabled",
            "method",
            "url",
            "body",
            "content_type",
            "body_mode",
            "timeout_ms",
            "expected_status",
            "key",
            "value",
            "comment",
        ]
    )
    for row in apis:
        enabled, name, method, url, body, content_type, body_mode, timeout_ms, expected, comment = row
        writer.writerow(
            ["API", name, enabled, method, url, body, content_type, body_mode, timeout_ms, expected, "", "", comment]
        )
    for enabled, api_name, key, value in queries:
        writer.writerow(["QUERY", api_name, enabled, "", "", "", "", "", "", "", key, value, ""])
    for enabled, api_name, key, value in headers:
        writer.writerow(["HEADER", api_name, enabled, "", "", "", "", "", "", "", key, value, ""])
    for enabled, api_name, key, value in forms:
        writer.writerow(["BODY", api_name, enabled, "", "", "", "", "", "", "", key, value, ""])
    for key, value, comment in settings:
        writer.writerow(["SETTING", "", "true", "", "", "", "", "", "", "", key, value, comment])
    return buffer.getvalue()


def template_sheets() -> List[Sheet]:
    tables = template_tables()
    widths = {
        "説明": [28, 88],
        "API": [12, 18, 12, 42, 36, 24, 14, 16, 16, 42],
        "クエリ": [12, 18, 18, 24],
        "ヘッダ": [12, 18, 22, 28],
        "フォーム": [12, 18, 18, 24],
        "設定": [24, 28, 56],
    }
    sheets: List[Sheet] = []
    for name, rows in tables.items():
        sheets.append(Sheet(name=name, rows=rows, widths=widths.get(name), wrap_cols=[2] if name == "説明" else None))
    return sheets


def write_templates(directory: str) -> Tuple[str, str]:
    os.makedirs(directory, exist_ok=True)
    csv_path = os.path.join(directory, "api-input-template.csv")
    xlsx_path = os.path.join(directory, "api-input-template.xlsx")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as handle:
        handle.write(template_csv_text())
    write_xlsx(xlsx_path, template_sheets())
    return csv_path, xlsx_path
