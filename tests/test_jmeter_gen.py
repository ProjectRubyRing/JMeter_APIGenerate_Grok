"""Tests for the JMeter 5.6.3 plan generator."""

from __future__ import annotations

import io
import os
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest
import zipfile
from xml.etree import ElementTree as ET

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.join(ROOT, "lib"))

from jmeter_gen.cli import main  # noqa: E402
from jmeter_gen.jmx import assert_jmx_balanced, build_jmx, display_url, resolve_api  # noqa: E402
from jmeter_gen.model import InputError  # noqa: E402
from jmeter_gen.parse import load_path, parse_csv_text, template_csv_text, write_templates  # noqa: E402
from jmeter_gen.xlsxio import read_xlsx, write_xlsx  # noqa: E402
from jmeter_gen.model import Sheet  # noqa: E402


def _jmx_facts(path):
    text = io.open(path, encoding="utf-8").read()
    assert_jmx_balanced(text)
    root = ET.fromstring(text)
    samplers = []
    for parent in root.iter("hashTree"):
        kids = list(parent)
        for index in range(0, len(kids) - 1, 2):
            element = kids[index]
            if element.tag != "HTTPSamplerProxy":
                continue
            tree = kids[index + 1]
            arguments = []
            for node in element.findall(".//elementProp[@elementType='HTTPArgument']"):
                name = node.find("./stringProp[@name='Argument.name']")
                value = node.find("./stringProp[@name='Argument.value']")
                arguments.append((None if name is None else name.text, None if value is None else value.text))
            headers = []
            for node in tree.findall(".//elementProp[@elementType='Header']"):
                headers.append(
                    (
                        node.find("./stringProp[@name='Header.name']").text,
                        node.find("./stringProp[@name='Header.value']").text,
                    )
                )
            status = tree.find(".//ResponseAssertion/collectionProp/stringProp")
            samplers.append(
                (
                    element.attrib["testname"],
                    element.find("./stringProp[@name='HTTPSampler.method']").text,
                    element.find("./stringProp[@name='HTTPSampler.protocol']").text,
                    element.find("./stringProp[@name='HTTPSampler.domain']").text,
                    element.find("./stringProp[@name='HTTPSampler.port']").text or "",
                    element.find("./stringProp[@name='HTTPSampler.path']").text,
                    element.find("./boolProp[@name='HTTPSampler.postBodyRaw']").text,
                    tuple(arguments),
                    tuple(headers),
                    None if status is None else status.text,
                    element.find("./stringProp[@name='HTTPSampler.response_timeout']").text or "",
                )
            )
    return (
        root.attrib["jmeter"],
        root.find(".//intProp[@name='ThreadGroup.num_threads']").text,
        root.find(".//stringProp[@name='LoopController.loops']").text,
        root.find(".//ResultCollector[@testname='Save JTL']//stringProp[@name='filename']").text,
        tuple(samplers),
    )


def _sampler_and_tree(root, name):
    for parent in root.iter("hashTree"):
        kids = list(parent)
        for index in range(0, len(kids) - 1, 2):
            element = kids[index]
            if element.tag == "HTTPSamplerProxy" and element.attrib.get("testname") == name:
                return element, kids[index + 1]
    raise AssertionError(name)


def fingerprint(plan):
    resolved = [resolve_api(api, plan.warnings) for api in plan.apis]
    rows = []
    for item in resolved:
        rows.append(
            (
                item.name,
                item.enabled,
                item.method,
                item.protocol,
                item.domain,
                item.port,
                item.path,
                tuple(item.arguments),
                item.raw_body,
                tuple(item.headers),
                item.expected_status,
                item.timeout_ms,
            )
        )
    return rows, plan.settings


class GeneratorTest(unittest.TestCase):
    def test_template_csv_and_xlsx_match(self):
        directory = tempfile.mkdtemp(prefix="jmeter-tpl-")
        try:
            csv_path, xlsx_path = write_templates(directory)
            with open(csv_path, "rb") as handle:
                self.assertTrue(handle.read().startswith(b"\xef\xbb\xbf"))
            csv_plan = load_path(csv_path)
            xlsx_plan = load_path(xlsx_path)
            self.assertEqual(fingerprint(csv_plan), fingerprint(xlsx_plan))
            settings = csv_plan.settings
            self.assertEqual(settings.threads, 1)
            self.assertEqual(settings.loops, 1)
            self.assertEqual(settings.ramp_up_sec, 1)
            enabled = [api for api in csv_plan.apis if api.enabled]
            self.assertEqual([api.method for api in enabled], ["GET", "POST", "POST"])
        finally:
            shutil.rmtree(directory)

    def test_jmx_contains_one_thread_and_listeners(self):
        plan = parse_csv_text(template_csv_text(), "template.csv")
        resolved = [resolve_api(api, plan.warnings) for api in plan.apis]
        xml_text = build_jmx(plan, resolved)
        assert_jmx_balanced(xml_text)
        root = ET.fromstring(xml_text)
        self.assertEqual(root.attrib["jmeter"], "5.6.3")
        self.assertEqual(root.attrib["version"], "1.2")
        threads = root.find(".//intProp[@name='ThreadGroup.num_threads']")
        loops = root.find(".//stringProp[@name='LoopController.loops']")
        self.assertEqual(threads.text, "1")
        self.assertEqual(loops.text, "1")
        names = [node.attrib["testname"] for node in root.findall(".//HTTPSamplerProxy")]
        self.assertEqual(names, ["get-item", "create-item", "submit-form"])
        get_item = root.findall(".//HTTPSamplerProxy")[0]
        self.assertEqual(get_item.find("./stringProp[@name='HTTPSampler.method']").text, "GET")
        self.assertEqual(get_item.find("./stringProp[@name='HTTPSampler.path']").text, "/v1/items")
        self.assertNotIn("?", get_item.find("./stringProp[@name='HTTPSampler.path']").text)
        argument_names = [
            node.find("./stringProp[@name='Argument.name']").text
            for node in get_item.findall(".//elementProp[@elementType='HTTPArgument']")
        ]
        self.assertEqual(argument_names, ["page", "size"])
        post, post_tree = _sampler_and_tree(root, "create-item")
        self.assertEqual(post.find("./boolProp[@name='HTTPSampler.postBodyRaw']").text, "true")
        body = post.find(".//stringProp[@name='Argument.value']").text
        self.assertIn('"name"', body)
        self.assertIn("sample", body)
        headers = [
            (
                node.find("./stringProp[@name='Header.name']").text,
                node.find("./stringProp[@name='Header.value']").text,
            )
            for node in post_tree.findall(".//elementProp[@elementType='Header']")
        ]
        self.assertIn(("Content-Type", "application/json"), headers)
        self.assertIn(("Accept", "application/json"), headers)
        form = root.findall(".//HTTPSamplerProxy")[2]
        self.assertEqual(form.find("./boolProp[@name='HTTPSampler.postBodyRaw']").text, "false")
        form_names = [
            node.find("./stringProp[@name='Argument.name']").text
            for node in form.findall(".//elementProp[@elementType='HTTPArgument']")
        ]
        self.assertEqual(form_names, ["title", "note"])
        jtl = root.find(".//ResultCollector[@testname='Save JTL']//stringProp[@name='filename']")
        self.assertEqual(jtl.text, "results/listener-results.jtl")
        script = root.find(".//JSR223Listener//stringProp[@name='script']")
        self.assertIn("API_RESULT", script.text)
        assertion = root.find(".//ResponseAssertion")
        self.assertIsNotNone(assertion)
        self.assertIn("200", ET.tostring(assertion, encoding="unicode"))

    def test_post_query_stays_on_path_and_xml_is_escaped(self):
        text = textwrap.dedent(
            """\
            record_type,api_name,enabled,method,url,body,content_type,body_mode,timeout_ms,expected_status,key,value,comment
            API,create,true,POST,https://api.example.com/v1/a?from=url,"a <b> &",,,5000,201,,,xml
            QUERY,create,true,,,,,,,,extra,1&2,,
            HEADER,create,true,,,,,,,,Accept,application/json,
            HEADER,create,true,,,,,,,,Authorization,Bearer secret-value,
            """
        )
        plan = parse_csv_text(text, "case.csv")
        resolved = [resolve_api(api, plan.warnings) for api in plan.apis]
        item = resolved[0]
        self.assertEqual(item.timeout_ms, "5000")
        self.assertIn("from=url", item.path)
        self.assertIn("extra=", item.path)
        self.assertNotIn("1&2", item.path)
        self.assertIn("1%261%262" if False else "1%262", item.path)
        self.assertIsNotNone(item.raw_body)
        xml_text = build_jmx(plan, resolved)
        self.assertNotIn("1&2", xml_text)
        self.assertIn("a &lt;b&gt; &amp;", xml_text)
        self.assertIn("Bearer secret-value", xml_text)
        self.assertIn("&quot;name&quot;" if False else "API_RESULT", xml_text)
        shown = display_url(item)
        self.assertTrue(shown.startswith("https://api.example.com/v1/a?"))

    def test_rejects_bad_method_duplicate_and_get_body(self):
        with self.assertRaises(InputError):
            parse_csv_text(
                "record_type,api_name,method,url\nAPI,one,PUT,https://example.com/\n",
                "bad.csv",
            )
        with self.assertRaises(InputError):
            parse_csv_text(
                "record_type,api_name,method,url\nAPI,one,GET,https://example.com/\nAPI,one,POST,https://example.com/\n",
                "dup.csv",
            )
        plan = parse_csv_text(
            "name,method,url,body\nget-one,GET,https://example.com/,{}\n",
            "get-body.csv",
        )
        with self.assertRaises(InputError):
            resolve_api(plan.apis[0], plan.warnings)

    def test_disabled_api_is_omitted_and_flat_csv_works(self):
        text = textwrap.dedent(
            """\
            name,method,url,query,headers,body,content_type
            keep,GET,https://example.com/keep,q=1,Accept: text/plain,,
            skip,POST,https://example.com/skip,,,{"a":1},application/json
            """
        )
        plan = parse_csv_text(text, "flat.csv")
        plan.apis[1].enabled = False
        resolved = [resolve_api(api, plan.warnings) for api in plan.apis]
        xml_text = build_jmx(plan, resolved)
        self.assertIn("keep", xml_text)
        self.assertNotIn("skip", xml_text)
        self.assertIn("text/plain", xml_text)

    def test_usage_masks_secrets_and_documents_both_os(self):
        directory = tempfile.mkdtemp(prefix="jmeter-out-")
        try:
            csv_path = os.path.join(directory, "in.csv")
            with open(csv_path, "w", encoding="utf-8", newline="\n") as handle:
                handle.write(
                    "record_type,api_name,enabled,method,url,key,value\n"
                    "API,demo,true,GET,https://example.com/demo,,\n"
                    "HEADER,demo,true,,,Authorization,Bearer secret-value\n"
                )
            output = os.path.join(directory, "out")
            code = main(["--input", csv_path, "--output-dir", output])
            self.assertEqual(code, 0)
            sheets = {sheet.name: sheet for sheet in read_xlsx(os.path.join(output, "jmeter-usage.xlsx"))}
            self.assertIn("CLI_RHEL9", sheets)
            self.assertIn("CLI_Windows11", sheets)
            joined = "\n".join(" | ".join(row) for row in sheets["生成されたAPI"].rows)
            self.assertIn("demo", joined)
            self.assertIn("********", joined)
            self.assertNotIn("secret-value", joined)
            jmx = open(os.path.join(output, "rest-api-test.jmx"), encoding="utf-8").read()
            self.assertIn("Bearer secret-value", jmx)
            self.assertIn('jmeter="5.6.3"', jmx)
        finally:
            shutil.rmtree(directory)

    def test_shared_strings_and_japanese_roundtrip(self):
        directory = tempfile.mkdtemp(prefix="jmeter-xlsx-")
        try:
            path = os.path.join(directory, "shared.xlsx")
            self._write_shared_string_workbook(path)
            sheets = read_xlsx(path)
            self.assertEqual(sheets[0].name, "API")
            self.assertEqual(sheets[0].rows[1][1], "取得")
            self.assertEqual(sheets[0].rows[1][3], "https://example.com/項目")
            plan = load_path(path)
            self.assertEqual(plan.apis[0].name, "取得")
            resolved = resolve_api(plan.apis[0], plan.warnings)
            self.assertEqual(resolved.path, "/%E9%A0%85%E7%9B%AE")
        finally:
            shutil.rmtree(directory)

    def test_cp932_csv_and_cli_override(self):
        directory = tempfile.mkdtemp(prefix="jmeter-cp932-")
        try:
            path = os.path.join(directory, "cp932.csv")
            text = "record_type,api_name,method,url,key,value\nAPI,項目,GET,https://example.com/a,,\nSETTING,,,,threads,4\n"
            with open(path, "wb") as handle:
                handle.write(text.encode("cp932"))
            output = os.path.join(directory, "out")
            code = main(["--input", path, "--output-dir", output, "--threads", "2", "--loops", "3"])
            self.assertEqual(code, 0)
            jmx = open(os.path.join(output, "rest-api-test.jmx"), encoding="utf-8").read()
            self.assertIn(">2</intProp>", jmx)
            self.assertIn(">3</stringProp>", jmx)
            self.assertIn("項目", jmx)
        finally:
            shutil.rmtree(directory)

    def test_shell_script_is_lf_and_runs(self):
        script = os.path.join(ROOT, "scripts", "generate-jmeter.sh")
        data = open(script, "rb").read()
        self.assertNotIn(b"\r", data)
        self.assertTrue(data.startswith(b"#!/usr/bin/env bash\n"))
        directory = os.path.join(ROOT, "out", "shell-test")
        if os.path.isdir(directory):
            shutil.rmtree(directory)
        os.makedirs(directory)
        try:
            csv_path = os.path.join(directory, "in.csv")
            with open(csv_path, "w", encoding="utf-8", newline="\n") as handle:
                handle.write(template_csv_text())
            output = os.path.join(directory, "out")
            bash = shutil.which("bash")
            self.assertIsNotNone(bash)
            completed = subprocess.run(
                [
                    bash,
                    "scripts/generate-jmeter.sh",
                    "--input",
                    "out/shell-test/in.csv",
                    "--output-dir",
                    "out/shell-test/out",
                ],
                cwd=ROOT,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                universal_newlines=True,
                check=False,
            )
            self.assertEqual(completed.returncode, 0, completed.stderr + completed.stdout)
            self.assertTrue(os.path.isfile(os.path.join(output, "rest-api-test.jmx")))
            self.assertTrue(os.path.isfile(os.path.join(output, "jmeter-usage.xlsx")))
        finally:
            shutil.rmtree(directory)

    def test_powershell_matches_python_on_template_and_shared_strings(self):
        directory = tempfile.mkdtemp(prefix="jmeter-ps-")
        try:
            csv_path, xlsx_path = write_templates(directory)
            shared_path = os.path.join(directory, "shared.xlsx")
            self._write_shared_string_workbook(shared_path)
            for name, source in (("csv", csv_path), ("xlsx", xlsx_path), ("shared", shared_path)):
                py_out = os.path.join(directory, "py-" + name)
                ps_out = os.path.join(directory, "ps-" + name)
                self.assertEqual(main(["--input", source, "--output-dir", py_out]), 0)
                completed = subprocess.run(
                    [
                        "powershell.exe",
                        "-NoProfile",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-File",
                        os.path.join(ROOT, "scripts", "generate-jmeter.ps1"),
                        "-InputPath",
                        source,
                        "-OutputDir",
                        ps_out,
                    ],
                    cwd=ROOT,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    check=False,
                )
                stdout = (completed.stdout or b"").decode("utf-8", errors="replace")
                stderr = (completed.stderr or b"").decode("cp932", errors="replace")
                self.assertEqual(completed.returncode, 0, stderr + "\n" + stdout)
                self.assertEqual(_jmx_facts(os.path.join(py_out, "rest-api-test.jmx")), _jmx_facts(os.path.join(ps_out, "rest-api-test.jmx")))
                py_usage = {sheet.name: sheet.rows for sheet in read_xlsx(os.path.join(py_out, "jmeter-usage.xlsx"))}
                ps_usage = {sheet.name: sheet.rows for sheet in read_xlsx(os.path.join(ps_out, "jmeter-usage.xlsx"))}
                self.assertEqual(py_usage["生成されたAPI"], ps_usage["生成されたAPI"])
                self.assertIn("CLI_RHEL9", ps_usage)
                self.assertIn("CLI_Windows11", ps_usage)
        finally:
            shutil.rmtree(directory)

    def test_unknown_child_is_rejected(self):
        text = "record_type,api_name,method,url,key,value\nQUERY,missing,,,,a,b\n"
        with self.assertRaises(InputError):
            parse_csv_text(text, "missing.csv")

    def _write_shared_string_workbook(self, path):
        shared = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
  <si><t>有効</t></si>
  <si><t>名前</t></si>
  <si><t>メソッド</t></si>
  <si><t>URL</t><rPh sb="0" eb="1"><t>ヨミ</t></rPh></si>
  <si><t>true</t></si>
  <si><r><t>取</t></r><r><t>得</t></r></si>
  <si><t>GET</t></si>
  <si><t>https://example.com/項目</t></si>
</sst>
"""
        sheet = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="s"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
      <c r="D1" t="s"><v>3</v></c>
    </row>
    <row r="3">
      <c r="A3" t="s"><v>4</v></c>
      <c r="B3" t="s"><v>5</v></c>
      <c r="C3" t="s"><v>6</v></c>
      <c r="D3" t="s"><v>7</v></c>
    </row>
  </sheetData>
</worksheet>
"""
        workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="API" sheetId="1" r:id="rId1"/></sheets>
</workbook>
"""
        rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>
"""
        root = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""
        content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
"""
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("[Content_Types].xml", content)
            archive.writestr("_rels/.rels", root)
            archive.writestr("xl/workbook.xml", workbook)
            archive.writestr("xl/_rels/workbook.xml.rels", rels)
            archive.writestr("xl/sharedStrings.xml", shared)
            archive.writestr("xl/worksheets/sheet1.xml", sheet)

    def test_writer_roundtrip_preserves_newline_and_markup(self):
        directory = tempfile.mkdtemp(prefix="jmeter-round-")
        try:
            path = os.path.join(directory, "book.xlsx")
            write_xlsx(
                path,
                [Sheet("API", [["名前", "値"], ["A & B", "line1\nline2"]], widths=[20, 20], wrap_cols=[2])],
            )
            sheets = read_xlsx(path)
            self.assertEqual(sheets[0].rows[1], ["A & B", "line1\nline2"])
        finally:
            shutil.rmtree(directory)


if __name__ == "__main__":
    unittest.main()
