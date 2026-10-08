"""Run Apache JMeter 5.6.3 in non-GUI mode."""

from __future__ import annotations

import csv
import os
import re
import shutil
import subprocess
import sys
from typing import List, Optional

from jmeter_gen.model import JMX_FILENAME, InputError


def _java_version_text() -> str:
    try:
        completed = subprocess.run(
            ["java", "-version"],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            universal_newlines=True,
            check=False,
        )
    except OSError as exc:
        raise InputError(
            "Java が見つかりません。JMeter 5.6.3 には Java 17 を推奨します。"
            " RHEL 9.8: sudo dnf install -y java-17-openjdk-headless"
        ) from exc
    return (completed.stderr or "") + (completed.stdout or "")


def ensure_java() -> None:
    text = _java_version_text()
    match = re.search(r'version "(?:1\.)?(\d+)', text)
    if match is None:
        raise InputError("java -version の結果を解釈できません。Java 17 をインストールしてください。")
    major = int(match.group(1))
    if major < 8 or major > 21:
        raise InputError(
            f"JMeter 5.6.3 は Java 8 から 21 で動作します。検出したメジャー版は {major} です。"
            " Java 17 を使ってください。"
        )


def find_jmeter(explicit: str) -> str:
    if explicit:
        if not os.path.isfile(explicit):
            raise InputError(f"JMeterの実行ファイルがありません: {explicit}")
        return os.path.abspath(explicit)
    candidates: List[str] = []
    home = os.environ.get("JMETER_HOME", "").strip()
    if home:
        if os.name == "nt":
            candidates.append(os.path.join(home, "bin", "jmeter.bat"))
        candidates.append(os.path.join(home, "bin", "jmeter"))
    for name in ("jmeter.bat", "jmeter") if os.name == "nt" else ("jmeter",):
        found = shutil.which(name)
        if found:
            candidates.append(found)
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            return os.path.abspath(candidate)
    raise InputError(
        "JMeter 5.6.3 が見つかりません。--jmeter か環境変数 JMETER_HOME を指定してください。"
    )


def summarize_jtl(path: str) -> dict:
    total = 0
    success = 0
    failed = 0
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        try:
            header = next(reader)
        except StopIteration:
            return {"total": 0, "success": 0, "failed": 0}
        try:
            index = [name.strip() for name in header].index("success")
        except ValueError:
            index = -1
        for row in reader:
            if not row or index < 0 or index >= len(row):
                continue
            total += 1
            if row[index].strip().lower() == "true":
                success += 1
            else:
                failed += 1
    return {"total": total, "success": success, "failed": failed}


def run_jmeter(binary: str, output_dir: str) -> int:
    ensure_java()
    results = os.path.join(output_dir, "results")
    os.makedirs(results, exist_ok=True)
    jtl = os.path.join(results, "results.jtl")
    log_path = os.path.join(results, "jmeter.log")
    report = os.path.join(results, "html-report")
    listener = os.path.join(results, "listener-results.jtl")
    for path in (jtl, log_path, listener):
        if os.path.isfile(path):
            os.remove(path)
    if os.path.isdir(report):
        shutil.rmtree(report)
    jmx = os.path.join(output_dir, JMX_FILENAME)
    properties = os.path.join(results, "saveservice.properties")
    with open(properties, "w", encoding="ascii", newline="\n") as handle:
        handle.write(
            "\n".join(
                [
                    "jmeter.save.saveservice.output_format=csv",
                    "jmeter.save.saveservice.print_field_names=true",
                    "jmeter.save.saveservice.default_delimiter=,",
                    "jmeter.save.saveservice.bytes=true",
                    "jmeter.save.saveservice.sent_bytes=true",
                    "jmeter.save.saveservice.label=true",
                    "jmeter.save.saveservice.latency=true",
                    "jmeter.save.saveservice.response_code=true",
                    "jmeter.save.saveservice.response_message=true",
                    "jmeter.save.saveservice.successful=true",
                    "jmeter.save.saveservice.thread_counts=true",
                    "jmeter.save.saveservice.thread_name=true",
                    "jmeter.save.saveservice.time=true",
                    "jmeter.save.saveservice.connect_time=true",
                    "jmeter.save.saveservice.url=true",
                    "jmeter.save.saveservice.timestamp_format=ms",
                    "",
                ]
            )
        )
    args = ["-n", "-q", properties, "-t", jmx, "-l", jtl, "-j", log_path, "-e", "-o", report]
    if binary.lower().endswith(".bat"):
        command = ["cmd.exe", "/d", "/c", subprocess.list2cmdline([binary] + args)]
    else:
        command = [binary] + args
    print("run: " + subprocess.list2cmdline(command))
    completed = subprocess.run(command, cwd=output_dir)
    if not os.path.isfile(jtl):
        print("error: JTLが出力されませんでした。jmeter.log を確認してください。", file=sys.stderr)
        return completed.returncode if completed.returncode else 1
    summary = summarize_jtl(jtl)
    print(f"samples: {summary['total']}")
    print(f"success: {summary['success']}")
    print(f"failed: {summary['failed']}")
    print(f"jtl: {jtl}")
    print(f"log: {log_path}")
    index = os.path.join(report, "index.html")
    if os.path.isfile(index):
        print(f"html: {index}")
    else:
        print("warning: HTMLレポートがありません。", file=sys.stderr)
    return completed.returncode
