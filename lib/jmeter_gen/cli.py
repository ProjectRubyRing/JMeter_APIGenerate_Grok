"""Command line entry for the JMeter plan generator."""

from __future__ import annotations

import argparse
import os
import sys
from typing import Optional, Sequence

from jmeter_gen.jmx import build_jmx, resolve_api
from jmeter_gen.model import JMX_FILENAME, USAGE_FILENAME, InputError, apply_overrides
from jmeter_gen.parse import load_path, write_templates
from jmeter_gen.runner import find_jmeter, run_jmeter
from jmeter_gen.usage import build_usage_sheets
from jmeter_gen.xlsxio import write_xlsx


def generate(input_path: str, output_dir: str, threads: Optional[int], loops: Optional[int], ramp_up: Optional[int]) -> str:
    plan = load_path(input_path)
    apply_overrides(plan.settings, threads, loops, ramp_up)
    resolved = [resolve_api(api, plan.warnings) for api in plan.apis]
    if not any(item.enabled for item in resolved):
        raise InputError("有効なAPIが1件もありません。")
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)
    jmx_path = os.path.join(output_dir, JMX_FILENAME)
    usage_path = os.path.join(output_dir, USAGE_FILENAME)
    jmx_text = build_jmx(plan, resolved)
    sheets = build_usage_sheets(plan, resolved, output_dir)
    temporary = jmx_path + ".tmp"
    with open(temporary, "w", encoding="utf-8", newline="\n") as handle:
        handle.write(jmx_text)
    os.replace(temporary, jmx_path)
    write_xlsx(usage_path, sheets)
    for warning in plan.warnings:
        print("warning: " + warning, file=sys.stderr)
    print("input: " + os.path.abspath(input_path))
    print("apis: " + str(len(plan.apis)))
    print("enabled: " + str(sum(1 for item in resolved if item.enabled)))
    print("threads: " + str(plan.settings.threads))
    print("loops: " + str(plan.settings.loops))
    print("jmx: " + jmx_path)
    print("usage: " + usage_path)
    return output_dir


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="generate-jmeter",
        description="Excel または CSV から JMeter 5.6.3 の JMX と、JMeterの利用手順Excelを生成します。",
    )
    parser.add_argument("-i", "--input", help="入力 .xlsx または .csv")
    parser.add_argument("-o", "--output-dir", help="JMX と手順Excelの出力ディレクトリ")
    parser.add_argument("--run", action="store_true", help="生成後に JMeter を非GUIで実行し、JTLとログとHTMLレポートを出力する")
    parser.add_argument("--jmeter", default="", help="jmeter または jmeter.bat のパス。未指定時は JMETER_HOME と PATH を見ます")
    parser.add_argument("--threads", type=int, default=None, help="スレッド数。未指定時は入力の設定、それも無ければ 1")
    parser.add_argument("--loops", type=int, default=None, help="ループ回数。未指定時は入力の設定、それも無ければ 1")
    parser.add_argument("--ramp-up", type=int, default=None, help="ランプアップ秒。未指定時は入力の設定、それも無ければ 1")
    parser.add_argument("--init-templates", action="store_true", help="入力テンプレート .xlsx と .csv を --output-dir に書き出す")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    try:
        if args.init_templates:
            if not args.output_dir:
                raise InputError("テンプレートの出力先を --output-dir で指定してください。")
            csv_path, xlsx_path = write_templates(args.output_dir)
            print("template_csv: " + os.path.abspath(csv_path))
            print("template_xlsx: " + os.path.abspath(xlsx_path))
            return 0
        if not args.input or not args.output_dir:
            parser.print_help(sys.stderr)
            return 2
        output_dir = generate(args.input, args.output_dir, args.threads, args.loops, args.ramp_up)
        if args.run:
            binary = find_jmeter(args.jmeter)
            return run_jmeter(binary, output_dir)
        return 0
    except InputError as exc:
        print("error: " + str(exc), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
