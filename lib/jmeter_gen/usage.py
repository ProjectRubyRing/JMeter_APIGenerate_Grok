"""Build the Excel workbook that explains how to run the generated plan."""

from __future__ import annotations

import os
from typing import List, Sequence

from jmeter_gen.jmx import display_url
from jmeter_gen.model import JMETER_VERSION, JMX_FILENAME, Plan, ResolvedHttp, Sheet, mask_secret


def _bash_quote(value: str) -> str:
    return "'" + value.replace("'", "'\\''") + "'"


def _ps_quote(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


def _pairs(rows: Sequence[Sequence[str]]) -> List[List[str]]:
    return [list(row) for row in rows]


def _command_block(output_dir: str) -> tuple:
    jmx = os.path.join(output_dir, JMX_FILENAME)
    results = os.path.join(output_dir, "results")
    jtl = os.path.join(results, "results.jtl")
    log_path = os.path.join(results, "jmeter.log")
    report = os.path.join(results, "html-report")
    listener = os.path.join(results, "listener-results.jtl")
    bash = "\n".join(
        [
            "cd " + _bash_quote(output_dir),
            "mkdir -p " + _bash_quote(results),
            "rm -f " + _bash_quote(jtl) + " " + _bash_quote(log_path) + " " + _bash_quote(listener),
            "rm -rf " + _bash_quote(report),
            '"$JMETER_HOME/bin/jmeter" -n \\',
            "  -t " + _bash_quote(jmx) + " \\",
            "  -l " + _bash_quote(jtl) + " \\",
            "  -j " + _bash_quote(log_path) + " \\",
            "  -e -o " + _bash_quote(report),
        ]
    )
    powershell = "\n".join(
        [
            "Set-Location -LiteralPath " + _ps_quote(output_dir),
            "New-Item -ItemType Directory -Force -Path " + _ps_quote(results) + " | Out-Null",
            "Remove-Item -LiteralPath " + _ps_quote(jtl) + "," + _ps_quote(log_path) + "," + _ps_quote(listener) + " -Force -ErrorAction SilentlyContinue",
            "if (Test-Path -LiteralPath " + _ps_quote(report) + ") { Remove-Item -LiteralPath " + _ps_quote(report) + " -Recurse -Force }",
            "& \"$env:JMETER_HOME\\bin\\jmeter.bat\" -n `",
            "  -t " + _ps_quote(jmx) + " `",
            "  -l " + _ps_quote(jtl) + " `",
            "  -j " + _ps_quote(log_path) + " `",
            "  -e -o " + _ps_quote(report),
        ]
    )
    return bash, powershell, jtl, log_path, report, listener


def build_usage_sheets(plan: Plan, resolved: Sequence[ResolvedHttp], output_dir: str) -> List[Sheet]:
    settings = plan.settings
    enabled = [item for item in resolved if item.enabled]
    bash, powershell, jtl, log_path, report, listener = _command_block(output_dir)
    overview = _pairs(
        [
            ["項目", "内容"],
            ["対象", f"Apache JMeter {JMETER_VERSION}"],
            ["テスト計画", settings.plan_name],
            ["JMX", os.path.join(output_dir, JMX_FILENAME)],
            ["入力", plan.source_path],
            ["スレッド数", str(settings.threads)],
            ["ループ回数", str(settings.loops)],
            ["ランプアップ秒", str(settings.ramp_up_sec)],
            ["呼び出し回数", "有効なAPIごとに、スレッド数 × ループ回数 です。初期値は各API 1回です。"],
            ["有効なAPI", str(len(enabled))],
            ["JTL（CLIの -l）", jtl],
            ["JTL（JMX内のリスナー）", listener + " 。作業ディレクトリが出力ディレクトリのとき、この相対パスに書かれます。"],
            ["ログ", log_path + " 。各サンプルは API_RESULT で検索できます。"],
            ["HTMLレポート", os.path.join(report, "index.html")],
            ["エラー時の動作", settings.on_error],
            ["リダイレクト追従", "true" if settings.follow_redirects else "false"],
            ["リスナーのJTL相対パス", settings.jtl_file],
        ]
    )
    environment = _pairs(
        [
            ["項目", "内容"],
            ["Java", "JMeter 5.6.3 は Java 8、11、17 で動作します。Java 17 を推奨します。Java 22 以降では動作しません。"],
            ["RHEL 9.8 の Java", "sudo dnf install -y java-17-openjdk-headless"],
            ["Windows 11 の Java", "Eclipse Temurin 17 などの JDK 17 を入れ、java が PATH で見えるようにします。"],
            ["JMeterの入手", "https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.zip を展開します。"],
            ["JMETER_HOME", "展開先の apache-jmeter-5.6.3 ディレクトリを JMETER_HOME にします。"],
            ["RHEL 9.8 の生成スクリプト", "bash と python3 が必要です。python3 が無い場合は sudo dnf install -y python3"],
            ["Windows 11 の生成スクリプト", "Windows PowerShell 5.1 だけで動作します。Excel や Python のインストールは不要です。"],
            ["追加のPythonパッケージ", "不要です。xlsx の読み書きは標準ライブラリだけで行います。"],
            ["秘密情報", "Authorization などのヘッダ値は JMX に平文で入ります。この手順書では伏せています。"],
        ]
    )
    gui = _pairs(
        [
            ["手順", "操作"],
            ["1", "RHELでは $JMETER_HOME/bin/jmeter 、Windowsでは %JMETER_HOME%\\bin\\jmeter.bat を起動します。"],
            ["2", "File、Open から生成された rest-api-test.jmx を開きます。"],
            ["3", "Thread Group REST API のスレッド数とループ回数が意図どおりか確認します。初期値はどちらも 1 です。"],
            ["4", "各 HTTP Request の Server Name、Port、Protocol、Path、Method、Parameters または Body Data、HTTP Headers を確認します。"],
            ["5", "起動前に、作業ディレクトリを出力ディレクトリにします。GUIのショートカットから起動した場合、JTLの相対パスは bin ディレクトリ基準になることがあります。"],
            ["6", "緑の Start を押します。初期値では有効な各APIへ 1 リクエストずつ送信します。"],
            ["7", "View Results Tree で要求と応答を確認します。Summary Report で成否と時間を確認します。"],
            ["8", "jmeter.log に API_RESULT で始まる行が出ます。GUI起動時のログは、多くの場合 JMeter の bin ディレクトリです。"],
            ["9", "Save JTL リスナーは " + settings.jtl_file + " にCSV形式のJTLを書きます。"],
        ]
    )
    rhel = _pairs(
        [
            ["項目", "内容"],
            ["前提", "export JMETER_HOME=/opt/apache-jmeter-5.6.3"],
            ["この出力向けのコマンド", bash],
            ["汎用の形", "\n".join([
                "cd \"$OUT\"",
                "mkdir -p \"$OUT/results\"",
                "rm -rf \"$OUT/results/html-report\"",
                "\"$JMETER_HOME/bin/jmeter\" -n \\",
                "  -t \"$OUT/rest-api-test.jmx\" \\",
                "  -l \"$OUT/results/results.jtl\" \\",
                "  -j \"$OUT/results/jmeter.log\" \\",
                "  -e -o \"$OUT/results/html-report\"",
            ])],
            ["生成と実行を一度に行う", "scripts/generate-jmeter.sh --input 入力.xlsx --output-dir 出力先 --run"],
            ["JMeterの指定", "--jmeter /opt/apache-jmeter-5.6.3/bin/jmeter または環境変数 JMETER_HOME"],
            ["作業ディレクトリ", "JMX内のJTLパスは相対パスです。必ず出力ディレクトリで jmeter を起動します。"],
        ]
    )
    windows = _pairs(
        [
            ["項目", "内容"],
            ["前提", "$env:JMETER_HOME = 'C:\\apache-jmeter-5.6.3'"],
            ["この出力向けのコマンド", powershell],
            ["汎用の形", "\n".join([
                "Set-Location -LiteralPath $Out",
                "New-Item -ItemType Directory -Force -Path (Join-Path $Out 'results') | Out-Null",
                "& \"$env:JMETER_HOME\\bin\\jmeter.bat\" -n `",
                "  -t (Join-Path $Out 'rest-api-test.jmx') `",
                "  -l (Join-Path $Out 'results\\results.jtl') `",
                "  -j (Join-Path $Out 'results\\jmeter.log') `",
                "  -e -o (Join-Path $Out 'results\\html-report')",
            ])],
            ["生成と実行を一度に行う", ".\\scripts\\generate-jmeter.ps1 -InputPath .\\入力.xlsx -OutputDir .\\出力先 -Run"],
            ["JMeterの指定", "-JMeterPath C:\\apache-jmeter-5.6.3\\bin\\jmeter.bat または環境変数 JMETER_HOME"],
            ["実行ポリシー", "powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\\scripts\\generate-jmeter.ps1 ..."],
        ]
    )
    results = _pairs(
        [
            ["項目", "内容"],
            ["results.jtl", "非GUIの -l で作るCSVです。HTMLレポートの入力になります。既存ファイルは追記されるため、再実行前に削除します。"],
            ["listener-results.jtl", "JMXの Simple Data Writer が書くCSVです。GUIでも非GUIでも、作業ディレクトリ基準で作られます。"],
            ["主要列", "timeStamp, elapsed, label, responseCode, responseMessage, threadName, success, bytes, sentBytes, URL, Latency, Connect"],
            ["success", "JMeterは応答コードだけでは失敗にしません。期待ステータスを指定したAPIだけ、コード不一致で success=false になります。"],
            ["jmeter.log", "起動ログと API_RESULT 行です。name、code、success、elapsed_ms、bytes、url を含みます。"],
            ["html-report", "jmeter の -e -o が作るダッシュボードです。index.html をブラウザで開きます。出力先ディレクトリは空である必要があります。"],
            ["再実行", "同じ -l ファイルが残っていると結果が混ざります。results.jtl、listener-results.jtl、jmeter.log、html-report を消してから実行します。"],
        ]
    )
    definition = _pairs(
        [
            ["項目", "内容"],
            ["Excel", "シート名は API、クエリ、ヘッダ、フォーム、設定です。説明シートは無視します。1行目がヘッダです。"],
            ["API列", "有効、名前、メソッド、URL、ボディ、Content-Type、ボディ形式、タイムアウトms、期待ステータス、コメント"],
            ["クエリ列", "有効、API名、名前、値。URLに元から付いているクエリも残します。"],
            ["ヘッダ列", "有効、API名、名前、値。"],
            ["フォーム列", "有効、API名、名前、値。POSTの application/x-www-form-urlencoded です。"],
            ["設定", "threads、loops、ramp_up_sec、follow_redirects、plan_name、on_error、jtl_file。初期値は threads=1、loops=1、ramp_up_sec=1 です。"],
            ["CSV", "UTF-8（BOM可）またはCP932。列は record_type,api_name,enabled,method,url,body,content_type,body_mode,timeout_ms,expected_status,key,value,comment"],
            ["CSVのrecord_type", "API、QUERY、HEADER、BODY、SETTING。日本語の クエリ、ヘッダ、フォーム、設定 も受け付けます。"],
            ["フラットCSV", "record_type が無く name,method,url がある場合です。query は key=value&key2=value2、headers は '名前: 値' を | で区切ります。"],
            ["メソッド", "GET と POST だけです。"],
            ["ボディ形式", "auto は、フォーム項目があればフォーム、JSONらしいボディなら raw と Content-Type を補います。"],
            ["POSTのクエリ", "POSTではクエリをURLのPathに付けます。Body Data とクエリが混ざらないようにするためです。"],
        ]
    )
    generated: List[List[str]] = [[
        "名前",
        "有効",
        "メソッド",
        "送信URL",
        "ボディ",
        "ヘッダ",
        "期待ステータス",
        "コメント",
    ]]
    for item in resolved:
        if item.raw_body is not None:
            body = item.raw_body
        elif item.method == "POST" and item.arguments:
            body = " | ".join(f"{key}={value}" for key, value in item.arguments)
        else:
            body = ""
        if len(body) > 4000:
            body = body[:4000] + " ...(省略)"
        header_text = " | ".join(f"{name}: {mask_secret(name, value)}" for name, value in item.headers)
        generated.append(
            [
                item.name,
                "true" if item.enabled else "false",
                item.method,
                display_url(item),
                body,
                header_text,
                item.expected_status,
                item.comment,
            ]
        )
    trouble = _pairs(
        [
            ["症状", "対処"],
            ["python3 が見つからない", "RHEL 9.8 で sudo dnf install -y python3 を実行します。"],
            ["Java が見つからない、または JMeter がすぐに止まる", "java -version が 17 であることを確認します。Java 22 以降は JMeter 5.6.3 では使えません。"],
            ["jmeter が見つからない", "--jmeter または -JMeterPath、あるいは JMETER_HOME を指定します。"],
            ["JTLに前回の結果が残る", "results.jtl は追記です。再実行前に削除します。--run と -Run は実行前に削除します。"],
            ["HTMLレポートが作られない", "html-report ディレクトリが残っていると失敗します。削除して再実行します。JTLはCSVである必要があります。"],
            ["相対パスのJTLが bin にできる", "出力ディレクトリへ移動してから jmeter を起動します。"],
            ["4xxなのに success が true", "期待ステータスを指定していないAPIは、応答を受け取れた時点で成功です。"],
            ["View Results Tree の警告", "非GUIでは表示用リスナーの警告が出ることがあります。1件の確認では結果ファイルには影響しません。"],
            ["POSTのJSONがクエリになる", "ボディ形式を raw にし、Content-Type を application/json にします。フォーム項目は空にします。"],
            ["ヘッダ名のエラー", "ヘッダ名は英数と記号だけです。値の改行は拒否します。"],
            ["証明書エラー", "JMeterは実行しているJavaの信頼ストアを使います。プライベートCAはJavaへ登録します。"],
        ]
    )
    return [
        Sheet("概要", overview, [28, 96], wrap_cols=[2]),
        Sheet("環境準備", environment, [28, 96], wrap_cols=[2]),
        Sheet("GUIでの実行", gui, [12, 96], wrap_cols=[2]),
        Sheet("CLI_RHEL9", rhel, [28, 110], wrap_cols=[2]),
        Sheet("CLI_Windows11", windows, [28, 110], wrap_cols=[2]),
        Sheet("結果の見方", results, [28, 96], wrap_cols=[2]),
        Sheet("入力定義", definition, [28, 110], wrap_cols=[2]),
        Sheet("生成されたAPI", generated, [18, 12, 12, 64, 42, 42, 16, 36], wrap_cols=[4, 5, 6, 8]),
        Sheet("トラブルシューティング", trouble, [36, 96], wrap_cols=[2]),
    ]
