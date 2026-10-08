"""Shared types and validation for the JMeter 5.6.3 plan generator."""

from __future__ import annotations

import os
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

JMETER_VERSION = "5.6.3"
DEFAULT_PLAN_NAME = "REST API Test Plan"
DEFAULT_JTL = "results/listener-results.jtl"
JMX_FILENAME = "rest-api-test.jmx"
USAGE_FILENAME = "jmeter-usage.xlsx"

_ILLEGAL_XML = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_HEADER_NAME = re.compile(r"^[!#$%&'*+\-.0-9A-Z^_`a-z|~]+$")
_SECRET_HEADER = re.compile(
    r"(authorization|api[-_]?key|token|secret|password|passwd|cookie|credential)",
    re.IGNORECASE,
)
_TRUE = {"1", "true", "yes", "y", "on", "はい", "有効", "○", "◯"}
_FALSE = {"0", "false", "no", "n", "off", "いいえ", "無効", "×", "x"}

ON_ERROR_VALUES = {
    "continue": "continue",
    "startnextloop": "startnextloop",
    "stopthread": "stopthread",
    "stoptest": "stoptest",
    "stoptestnow": "stoptestnow",
    "続行": "continue",
    "停止": "stoptest",
}


class InputError(Exception):
    """The API definition or the command line cannot be used."""


def normalize_token(value: object) -> str:
    text = unicodedata.normalize("NFKC", "" if value is None else str(value))
    text = text.strip().lower()
    text = text.replace("-", "_").replace(" ", "").replace("\u3000", "")
    return text


def require_xml_text(value: object, label: str) -> str:
    text = "" if value is None else str(value)
    if _ILLEGAL_XML.search(text):
        raise InputError(f"{label} に使用できない制御文字が含まれています。")
    return text


def xml_escape(value: object) -> str:
    text = require_xml_text(value, "値")
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def parse_bool(value: object, default: bool = True, label: str = "有効") -> bool:
    if value is None:
        return default
    text = str(value).strip()
    if text == "":
        return default
    token = normalize_token(text)
    if token in _TRUE:
        return True
    if token in _FALSE:
        return False
    raise InputError(f"{label} の値を解釈できません: {value}")


def parse_int_range(value: object, label: str, minimum: int, maximum: int) -> int:
    text = str(value).strip()
    if not re.fullmatch(r"[0-9]+", text):
        raise InputError(f"{label} は整数で指定してください: {value}")
    number = int(text)
    if number < minimum or number > maximum:
        raise InputError(
            f"{label} は {minimum} 以上 {maximum} 以下にしてください: {number}"
        )
    return number


def mask_secret(name: str, value: str) -> str:
    if _SECRET_HEADER.search(name or ""):
        return "********"
    return value


def validate_header_name(name: object) -> str:
    text = "" if name is None else str(name).strip()
    if not text or _HEADER_NAME.fullmatch(text) is None:
        raise InputError(f"ヘッダ名が不正です: {name}")
    return text


def validate_header_value(value: object) -> str:
    text = "" if value is None else str(value).strip()
    if re.search(r"[\r\n]", text):
        raise InputError("ヘッダ値に改行は使えません。")
    require_xml_text(text, "ヘッダ値")
    return text


def validate_param_name(name: object, label: str) -> str:
    text = "" if name is None else str(name).strip()
    if text == "":
        raise InputError(f"{label} が空です。")
    if re.search(r"[\r\n]", text):
        raise InputError(f"{label} に改行は使えません。")
    require_xml_text(text, label)
    return text


def validate_relative_jtl(value: object) -> str:
    text = "" if value is None else str(value).strip()
    if not text:
        raise InputError("jtl_file が空です。")
    if os.path.isabs(text) or re.match(r"^[A-Za-z]:", text):
        raise InputError(
            f"jtl_file は出力ディレクトリ内の相対パスにしてください: {text}"
        )
    parts = re.split(r"[\\/]", text)
    if any(part in ("", "..") for part in parts):
        raise InputError(f"jtl_file が不正です: {text}")
    require_xml_text(text, "jtl_file")
    return text.replace("\\", "/")


@dataclass
class ApiRequest:
    name: str
    method: str
    url: str
    body: str = ""
    content_type: str = ""
    body_mode: str = "auto"
    timeout_ms: str = ""
    expected_status: str = ""
    comment: str = ""
    enabled: bool = True
    query: List[Tuple[str, str]] = field(default_factory=list)
    headers: List[Tuple[str, str]] = field(default_factory=list)
    form_params: List[Tuple[str, str]] = field(default_factory=list)
    row_label: str = ""


@dataclass
class Settings:
    threads: int = 1
    loops: int = 1
    ramp_up_sec: int = 1
    follow_redirects: bool = True
    plan_name: str = DEFAULT_PLAN_NAME
    on_error: str = "continue"
    jtl_file: str = DEFAULT_JTL


@dataclass
class Plan:
    apis: List[ApiRequest]
    settings: Settings
    warnings: List[str]
    source_path: str


@dataclass
class ResolvedHttp:
    name: str
    method: str
    protocol: str
    domain: str
    port: str
    path: str
    arguments: List[Tuple[str, str]]
    raw_body: Optional[str]
    headers: List[Tuple[str, str]]
    timeout_ms: str
    expected_status: str
    comment: str
    original_url: str
    enabled: bool


@dataclass
class Sheet:
    name: str
    rows: List[List[str]]
    row_numbers: Optional[List[int]] = None
    widths: Optional[List[float]] = None
    wrap_cols: Optional[Sequence[int]] = None


def apply_overrides(
    settings: Settings,
    threads: Optional[int],
    loops: Optional[int],
    ramp_up: Optional[int],
) -> None:
    if threads is not None:
        settings.threads = parse_int_range(threads, "threads", 1, 1000000)
    if loops is not None:
        settings.loops = parse_int_range(loops, "loops", 1, 1000000)
    if ramp_up is not None:
        settings.ramp_up_sec = parse_int_range(ramp_up, "ramp-up", 0, 86400)
