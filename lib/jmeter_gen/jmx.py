"""Build an Apache JMeter 5.6.3 test plan."""

from __future__ import annotations

import re
from typing import List, Optional, Sequence, Tuple
from urllib.parse import quote, urlsplit
from xml.etree import ElementTree as ET

from jmeter_gen.model import (
    JMETER_VERSION,
    ApiRequest,
    InputError,
    Plan,
    ResolvedHttp,
    normalize_token,
    parse_int_range,
    require_xml_text,
    xml_escape,
)
from jmeter_gen.parse import parse_qsl_strict

_BODY_MODES = {
    "auto": "auto",
    "自動": "auto",
    "raw": "raw",
    "json": "raw",
    "生": "raw",
    "そのまま": "raw",
    "form": "form",
    "フォーム": "form",
    "パラメータ": "form",
}

_GROOVY = (
    'def urlText = ""; '
    'try { urlText = String.valueOf(prev.getURL()) } catch (Exception ignored) { urlText = "" }; '
    'log.info("API_RESULT name=" + prev.getSampleLabel()'
    ' + " code=" + prev.getResponseCode()'
    ' + " success=" + prev.isSuccessful()'
    ' + " elapsed_ms=" + prev.getTime()'
    ' + " bytes=" + prev.getBytesAsLong()'
    ' + " url=" + urlText)'
)


def _mode(api: ApiRequest) -> str:
    token = normalize_token(api.body_mode) if api.body_mode.strip() else "auto"
    if token not in _BODY_MODES:
        raise InputError(f"{api.name}: ボディ形式が不正です: {api.body_mode}")
    mode = _BODY_MODES[token]
    has_form = bool(api.form_params)
    has_body = api.body != ""
    content_type = (api.content_type or "").lower()
    if "multipart/form-data" in content_type:
        raise InputError(f"{api.name}: multipart/form-data には対応していません。")
    if api.method == "GET" and (has_body or has_form or mode in ("raw", "form")):
        raise InputError(f"{api.name}: GET にリクエストボディは指定できません。")
    if mode == "raw":
        if has_form:
            raise InputError(f"{api.name}: ボディ形式 raw とフォーム項目は同時に指定できません。")
        return "raw" if has_body else "empty"
    if mode == "form":
        if has_body and has_form:
            raise InputError(f"{api.name}: リクエストボディとフォーム項目は同時に指定できません。")
        if has_body:
            return "form_body"
        if has_form:
            return "form_params"
        raise InputError(f"{api.name}: フォーム項目がありません。")
    if has_form and has_body:
        raise InputError(f"{api.name}: リクエストボディとフォーム項目は同時に指定できません。")
    if has_form:
        if "json" in content_type:
            raise InputError(f"{api.name}: Content-Type が JSON ですがフォーム項目があります。")
        return "form_params"
    if not has_body:
        return "empty"
    if "application/x-www-form-urlencoded" in content_type:
        return "form_body"
    return "raw"


def _split_url(api: ApiRequest, warnings: List[str]) -> Tuple[str, str, str, str, List[Tuple[str, str]]]:
    try:
        parts = urlsplit(api.url.strip())
    except ValueError as exc:
        raise InputError(f"{api.name}: URLが不正です: {api.url}") from exc
    scheme = parts.scheme.lower()
    if scheme not in ("http", "https"):
        raise InputError(f"{api.name}: URLは http または https にしてください: {api.url}")
    if parts.username or parts.password:
        raise InputError(f"{api.name}: URLにユーザー情報を含めないでください。認証はヘッダで指定します。")
    if parts.fragment:
        warnings.append(f"{api.name}: URLのフラグメントは送信されないため無視しました。")
    host = parts.hostname
    if not host:
        raise InputError(f"{api.name}: URLにホスト名がありません: {api.url}")
    try:
        port = parts.port
    except ValueError as exc:
        raise InputError(f"{api.name}: ポートが不正です: {api.url}") from exc
    port_text = ""
    if port is not None:
        if port < 1 or port > 65535:
            raise InputError(f"{api.name}: ポートが不正です: {port}")
        default_port = 443 if scheme == "https" else 80
        if port != default_port:
            port_text = str(port)
    domain = f"[{host}]" if ":" in host else host
    path = parts.path or "/"
    try:
        query = parse_qsl_strict(parts.query) if parts.query else []
    except ValueError as exc:
        raise InputError(f"{api.name}: URLのクエリが不正です。") from exc
    cleaned: List[Tuple[str, str]] = []
    for key, value in query:
        cleaned.append((require_xml_text(key, f"{api.name} のクエリ名"), require_xml_text(value, f"{api.name} のクエリ値")))
        if key == "":
            raise InputError(f"{api.name}: URLのクエリ名が空です。")
    return scheme, domain, port_text, path, cleaned


def _encode_path(path: str) -> str:
    if not path.startswith("/"):
        path = "/" + path
    return quote(path, safe="/%:@-._~!$&'()*+,;=")


def _append_query(path: str, pairs: Sequence[Tuple[str, str]]) -> str:
    if not pairs:
        return path
    encoded = "&".join(
        quote(key, safe="-._~") + "=" + quote(value, safe="-._~") for key, value in pairs
    )
    return path + "?" + encoded


def _merge_headers(api: ApiRequest, mode: str) -> List[Tuple[str, str]]:
    headers = list(api.headers)
    content_indexes = [index for index, (name, _) in enumerate(headers) if name.lower() == "content-type"]
    if len(content_indexes) > 1:
        raise InputError(f"{api.name}: Content-Type ヘッダが重複しています。")
    user_type = (api.content_type or "").strip()
    if user_type and content_indexes:
        current = headers[content_indexes[0]][1].strip()
        if current.lower() != user_type.lower():
            raise InputError(f"{api.name}: Content-Type 列とヘッダの値が矛盾しています。")
        return headers
    if user_type:
        headers.insert(0, ("Content-Type", user_type))
        return headers
    if content_indexes:
        return headers
    if mode == "raw":
        sniff = api.body.lstrip("\ufeff \t\r\n")
        if sniff.startswith("{") or sniff.startswith("["):
            headers.insert(0, ("Content-Type", "application/json; charset=UTF-8"))
    elif mode in ("form_body", "form_params"):
        headers.insert(0, ("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"))
    return headers


def _timeout(api: ApiRequest) -> str:
    if str(api.timeout_ms).strip() == "":
        return ""
    return str(parse_int_range(api.timeout_ms, f"{api.name} のタイムアウト", 1, 3600000))


def _expected_status(api: ApiRequest) -> str:
    status = str(api.expected_status).strip()
    if status == "":
        return ""
    if re.fullmatch(r"[1-5][0-9][0-9]", status) is None:
        raise InputError(f"{api.name}: 期待ステータスは 100 から 599 の整数です: {status}")
    return status


def resolve_api(api: ApiRequest, warnings: List[str]) -> ResolvedHttp:
    mode = _mode(api)
    scheme, domain, port, raw_path, url_query = _split_url(api, warnings)
    path = _encode_path(raw_path)
    query_pairs: List[Tuple[str, str]] = []
    for key, value in list(url_query) + list(api.query):
        query_pairs.append((require_xml_text(key, f"{api.name} のクエリ名"), require_xml_text(value, f"{api.name} のクエリ値")))
        if key == "":
            raise InputError(f"{api.name}: クエリ名が空です。")
    headers = _merge_headers(api, mode)
    seen = set()
    for name, _ in headers:
        lowered = name.lower()
        if lowered in seen and lowered != "content-type":
            warnings.append(f"{api.name}: ヘッダが重複しています: {name}")
        seen.add(lowered)
    arguments: List[Tuple[str, str]] = []
    raw_body: Optional[str] = None
    if api.method == "GET":
        arguments = query_pairs
        final_path = path
    elif mode == "form_body":
        try:
            arguments = parse_qsl_strict(api.body)
        except ValueError as exc:
            raise InputError(f"{api.name}: フォームボディは key=value&key2=value2 の形式にしてください。") from exc
        if not arguments:
            raise InputError(f"{api.name}: フォーム項目がありません。")
        arguments = [(require_xml_text(key, f"{api.name} のフォーム名"), require_xml_text(value, f"{api.name} のフォーム値")) for key, value in arguments]
        final_path = _append_query(path, query_pairs)
    elif mode == "form_params":
        arguments = [(require_xml_text(key, f"{api.name} のフォーム名"), require_xml_text(value, f"{api.name} のフォーム値")) for key, value in api.form_params]
        final_path = _append_query(path, query_pairs)
    elif mode == "raw":
        raw_body = api.body
        final_path = _append_query(path, query_pairs)
    else:
        final_path = _append_query(path, query_pairs)
    return ResolvedHttp(
        name=api.name,
        method=api.method,
        protocol=scheme,
        domain=domain,
        port=port,
        path=final_path,
        arguments=arguments,
        raw_body=raw_body,
        headers=headers,
        timeout_ms=_timeout(api),
        expected_status=_expected_status(api),
        comment=api.comment,
        original_url=api.url.strip(),
        enabled=api.enabled,
    )


def display_url(item: ResolvedHttp) -> str:
    port = f":{item.port}" if item.port else ""
    base = f"{item.protocol}://{item.domain}{port}{item.path}"
    if item.method == "GET" and item.arguments and "?" not in item.path:
        query = "&".join(
            quote(key, safe="-._~") + "=" + quote(value, safe="-._~") for key, value in item.arguments
        )
        return base + "?" + query
    return base


def _save_config(include_response: bool) -> str:
    response = "true" if include_response else "false"
    return f"""<objProp>
              <name>saveConfig</name>
              <value class="SampleSaveConfiguration">
                <time>true</time>
                <latency>true</latency>
                <timestamp>true</timestamp>
                <success>true</success>
                <label>true</label>
                <code>true</code>
                <message>true</message>
                <threadName>true</threadName>
                <dataType>true</dataType>
                <encoding>false</encoding>
                <assertions>true</assertions>
                <subresults>true</subresults>
                <responseData>{response}</responseData>
                <samplerData>false</samplerData>
                <xml>false</xml>
                <fieldNames>true</fieldNames>
                <responseHeaders>false</responseHeaders>
                <requestHeaders>false</requestHeaders>
                <responseDataOnError>false</responseDataOnError>
                <saveAssertionResultsFailureMessage>true</saveAssertionResultsFailureMessage>
                <assertionsResultsToSave>0</assertionsResultsToSave>
                <bytes>true</bytes>
                <sentBytes>true</sentBytes>
                <url>true</url>
                <threadCounts>true</threadCounts>
                <idleTime>true</idleTime>
                <connectTime>true</connectTime>
              </value>
            </objProp>"""


def _string_prop(name: str, value: str) -> str:
    return f'<stringProp name="{xml_escape(name)}">{xml_escape(value)}</stringProp>'


def _bool_prop(name: str, value: bool) -> str:
    return f'<boolProp name="{xml_escape(name)}">{str(value).lower()}</boolProp>'


def _argument_xml(name: Optional[str], value: str, encode: bool) -> str:
    name_attr = "" if name is None else xml_escape(name)
    lines = [
        f'<elementProp name="{name_attr}" elementType="HTTPArgument">',
        _bool_prop("HTTPArgument.always_encode", encode),
        _string_prop("Argument.value", value),
        _string_prop("Argument.metadata", "="),
        _bool_prop("HTTPArgument.use_equals", True),
    ]
    if name is not None:
        lines.append(_string_prop("Argument.name", name))
    lines.append("</elementProp>")
    return "\n".join(lines)


def _sampler_xml(item: ResolvedHttp, follow_redirects: bool) -> str:
    if item.raw_body is None:
        arguments = "\n".join(_argument_xml(name, value, True) for name, value in item.arguments)
        post_raw = False
    else:
        arguments = _argument_xml(None, item.raw_body, False)
        post_raw = True
    argument_block = f"""<elementProp name="HTTPsampler.Arguments" elementType="Arguments" guiclass="HTTPArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
          <collectionProp name="Arguments.arguments">
            {arguments}
          </collectionProp>
        </elementProp>"""
    return f"""<HTTPSamplerProxy guiclass="HttpTestSampleGui" testclass="HTTPSamplerProxy" testname="{xml_escape(item.name)}" enabled="true">
        {_bool_prop("HTTPSampler.postBodyRaw", post_raw)}
        {argument_block}
        {_string_prop("HTTPSampler.domain", item.domain)}
        {_string_prop("HTTPSampler.port", item.port)}
        {_string_prop("HTTPSampler.protocol", item.protocol)}
        {_string_prop("HTTPSampler.contentEncoding", "UTF-8")}
        {_string_prop("HTTPSampler.path", item.path)}
        {_string_prop("HTTPSampler.method", item.method)}
        {_bool_prop("HTTPSampler.follow_redirects", follow_redirects)}
        {_bool_prop("HTTPSampler.auto_redirects", False)}
        {_bool_prop("HTTPSampler.use_keepalive", True)}
        {_bool_prop("HTTPSampler.DO_MULTIPART_POST", False)}
        {_string_prop("HTTPSampler.embedded_url_re", "")}
        {_string_prop("HTTPSampler.connect_timeout", item.timeout_ms)}
        {_string_prop("HTTPSampler.response_timeout", item.timeout_ms)}
        {_string_prop("HTTPSampler.implementation", "HttpClient4")}
      </HTTPSamplerProxy>"""


def _header_manager(item: ResolvedHttp) -> str:
    if not item.headers:
        return ""
    headers = "\n".join(
        "<elementProp name=\"\" elementType=\"Header\">\n"
        f"{_string_prop('Header.name', name)}\n"
        f"{_string_prop('Header.value', value)}\n"
        "</elementProp>"
        for name, value in item.headers
    )
    return f"""<HeaderManager guiclass="HeaderPanel" testclass="HeaderManager" testname="HTTP Headers" enabled="true">
            <collectionProp name="HeaderManager.headers">
              {headers}
            </collectionProp>
          </HeaderManager>
          <hashTree/>"""


def _assertion(item: ResolvedHttp) -> str:
    if not item.expected_status:
        return ""
    return f"""<ResponseAssertion guiclass="AssertionGui" testclass="ResponseAssertion" testname="Expected status {xml_escape(item.expected_status)}" enabled="true">
            <collectionProp name="Asserion.test_strings">
              <stringProp name="0">{xml_escape(item.expected_status)}</stringProp>
            </collectionProp>
            {_string_prop("Assertion.custom_message", "Expected HTTP status " + item.expected_status)}
            {_string_prop("Assertion.test_field", "Assertion.response_code")}
            {_bool_prop("Assertion.assume_success", False)}
            <intProp name="Assertion.test_type">8</intProp>
          </ResponseAssertion>
          <hashTree/>"""


def _result_collector(gui: str, title: str, filename: str, include_response: bool) -> str:
    return f"""<ResultCollector guiclass="{gui}" testclass="ResultCollector" testname="{xml_escape(title)}" enabled="true">
          {_bool_prop("ResultCollector.error_logging", False)}
          {_save_config(include_response)}
          {_string_prop("filename", filename)}
        </ResultCollector>
        <hashTree/>"""


def _listener_xml() -> str:
    return f"""<JSR223Listener guiclass="TestBeanGUI" testclass="JSR223Listener" testname="Log sample result" enabled="true">
          {_string_prop("cacheKey", "true")}
          {_string_prop("filename", "")}
          {_string_prop("parameters", "")}
          {_string_prop("script", _GROOVY)}
          {_string_prop("scriptLanguage", "groovy")}
        </JSR223Listener>
        <hashTree/>"""


def assert_jmx_balanced(xml_text: str) -> None:
    root = ET.fromstring(xml_text)
    if root.tag != "jmeterTestPlan":
        raise RuntimeError("JMXのルート要素が不正です。")
    children = list(root)
    if len(children) != 1 or children[0].tag != "hashTree":
        raise RuntimeError("JMXのhashTreeが不正です。")

    def check(node: ET.Element) -> None:
        kids = list(node)
        if len(kids) % 2 != 0:
            raise RuntimeError("JMXの要素とhashTreeが対になっていません。")
        for index in range(0, len(kids), 2):
            element, tree = kids[index], kids[index + 1]
            if element.tag == "hashTree" or tree.tag != "hashTree":
                raise RuntimeError("JMXのhashTreeの並びが不正です。")
            check(tree)

    check(children[0])


def build_jmx(plan: Plan, resolved: Sequence[ResolvedHttp]) -> str:
    enabled = [item for item in resolved if item.enabled]
    if not enabled:
        raise InputError("有効なAPIが1件もありません。")
    settings = plan.settings
    comments = (
        f"Apache JMeter {JMETER_VERSION} 向けに自動生成したREST APIテストです。"
        f"スレッド数={settings.threads}、ループ回数={settings.loops}、ランプアップ={settings.ramp_up_sec}秒。"
        "初期値では有効な各APIを1回呼び出します。"
        f"Simple Data Writer は作業ディレクトリ基準の {settings.jtl_file} にJTLを書きます。"
        "JSR223 Listener は jmeter.log に API_RESULT 行を出します。"
        "非GUIで実行するときは、出力ディレクトリを作業ディレクトリにしてください。"
    )
    samplers = []
    for item in enabled:
        children = "\n".join(part for part in (_header_manager(item), _assertion(item)) if part)
        samplers.append(
            _sampler_xml(item, settings.follow_redirects)
            + "\n      <hashTree>\n"
            + children
            + "\n      </hashTree>"
        )
    listeners = "\n".join(
        [
            _listener_xml(),
            _result_collector("SimpleDataWriter", "Save JTL", settings.jtl_file, False),
            _result_collector("SummaryReport", "Summary Report", "", False),
            _result_collector("ViewResultsFullVisualizer", "View Results Tree", "", True),
        ]
    )
    xml_text = f"""<?xml version="1.0" encoding="UTF-8"?>
<jmeterTestPlan version="1.2" properties="5.0" jmeter="{JMETER_VERSION}">
  <hashTree>
    <TestPlan guiclass="TestPlanGui" testclass="TestPlan" testname="{xml_escape(settings.plan_name)}" enabled="true">
      {_string_prop("TestPlan.comments", comments)}
      {_bool_prop("TestPlan.functional_mode", False)}
      {_bool_prop("TestPlan.tearDown_on_shutdown", True)}
      {_bool_prop("TestPlan.serialize_threadgroups", False)}
      <elementProp name="TestPlan.user_defined_variables" elementType="Arguments" guiclass="ArgumentsPanel" testclass="Arguments" testname="User Defined Variables" enabled="true">
        <collectionProp name="Arguments.arguments"/>
      </elementProp>
      {_string_prop("TestPlan.user_define_classpath", "")}
    </TestPlan>
    <hashTree>
      <ThreadGroup guiclass="ThreadGroupGui" testclass="ThreadGroup" testname="REST API" enabled="true">
        {_string_prop("ThreadGroup.on_sample_error", settings.on_error)}
        <elementProp name="ThreadGroup.main_controller" elementType="LoopController" guiclass="LoopControlPanel" testclass="LoopController" testname="Loop Controller" enabled="true">
          {_bool_prop("LoopController.continue_forever", False)}
          {_string_prop("LoopController.loops", str(settings.loops))}
        </elementProp>
        <intProp name="ThreadGroup.num_threads">{settings.threads}</intProp>
        <intProp name="ThreadGroup.ramp_time">{settings.ramp_up_sec}</intProp>
        {_bool_prop("ThreadGroup.scheduler", False)}
        {_string_prop("ThreadGroup.duration", "")}
        {_string_prop("ThreadGroup.delay", "")}
        {_bool_prop("ThreadGroup.same_user_on_next_iteration", True)}
      </ThreadGroup>
      <hashTree>
      {chr(10).join(samplers)}
      {listeners}
      </hashTree>
    </hashTree>
  </hashTree>
</jmeterTestPlan>
"""
    assert_jmx_balanced(xml_text)
    return xml_text
