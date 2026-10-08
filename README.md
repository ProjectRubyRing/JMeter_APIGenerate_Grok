# JMeter REST API テスト計画の生成

Excel または CSV に書いた REST API 定義から、Apache JMeter 5.6.3 のテスト計画（`.jmx`）と、JMeter の利用手順（`.xlsx`）を作ります。

初期値はスレッド 1、ループ 1 です。有効な各 API を 1 回呼び出します。生成した JMX は、実行時に CSV 形式の JTL と `jmeter.log` の `API_RESULT` 行へ結果を出します。

## 必要なもの

| 環境 | 生成 | テスト実行 |
| --- | --- | --- |
| RHEL 9.8 | bash と python3（標準ライブラリのみ。`sudo dnf install -y python3`） | Java 17 と JMeter 5.6.3 |
| Windows 11 | Windows PowerShell 5.1（Excel も Python も不要） | Java 17 と JMeter 5.6.3 |

JMeter 5.6.3 は Java 8、11、17 で動作します。Java 22 以降では動作しません。Java 17 を推奨します。

```text
https://archive.apache.org/dist/jmeter/binaries/apache-jmeter-5.6.3.zip
```

展開先を `JMETER_HOME` にします。

## 入力

テンプレートは `templates/api-input-template.xlsx` と `templates/api-input-template.csv` です。再作成するときは次のどちらかです。

```bash
./scripts/generate-jmeter.sh --init-templates --output-dir templates
```

```powershell
.\scripts\generate-jmeter.ps1 -InitTemplates -OutputDir .\templates
```

### Excel

1 行目がヘッダです。説明シートは読み込みません。

| シート | 内容 |
| --- | --- |
| API | 有効、名前、メソッド、URL、ボディ、Content-Type、ボディ形式、タイムアウトms、期待ステータス、コメント |
| クエリ | 有効、API名、名前、値 |
| ヘッダ | 有効、API名、名前、値 |
| フォーム | 有効、API名、名前、値。`application/x-www-form-urlencoded` |
| 設定 | threads、loops、ramp_up_sec、follow_redirects、plan_name、on_error、jtl_file |

メソッドは `GET` と `POST` です。ボディ形式が空なら自動判定です。`raw` は JSON などの生ボディ、`form` はフォーム項目です。POST のクエリは URL のパスに付け、ボディと混ざりません。

### CSV

UTF-8（BOM 可）または CP932 です。列は次です。

```text
record_type,api_name,enabled,method,url,body,content_type,body_mode,timeout_ms,expected_status,key,value,comment
```

`record_type` は `API`、`QUERY`、`HEADER`、`BODY`、`SETTING` です。`record_type` が無く `name,method,url` があるフラット形式も読めます。フラット形式の `query` と `form` は `key=value&key2=value2`、`headers` は `名前: 値` を `|` で区切ります。

## 生成

RHEL 9.8:

```bash
chmod +x scripts/generate-jmeter.sh
./scripts/generate-jmeter.sh \
  --input templates/api-input-template.xlsx \
  --output-dir out
```

Windows 11:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\generate-jmeter.ps1 `
  -InputPath .\templates\api-input-template.xlsx `
  -OutputDir .\out
```

出力:

- `out/rest-api-test.jmx`
- `out/jmeter-usage.xlsx`（GUI と CLI の手順、RHEL と Windows のコマンド、結果の見方、今回の API 一覧）

スレッド数とループ回数は、入力の設定、無ければ 1 です。コマンドの `--threads` / `-Threads` と `--loops` / `-Loops` が優先されます。

## 実行

生成だけ行うのが初期動作です。JMeter で送るときは `--run` または `-Run` を付けます。作業ディレクトリは出力ディレクトリになり、次のファイルができます。

- `results/results.jtl`（非 GUI の `-l`。HTML レポートの入力）
- `results/listener-results.jtl`（JMX 内の Simple Data Writer）
- `results/jmeter.log`（`API_RESULT` 行）
- `results/html-report/index.html`

```bash
export JMETER_HOME=/opt/apache-jmeter-5.6.3
./scripts/generate-jmeter.sh --input templates/api-input-template.csv --output-dir out --run
```

```powershell
$env:JMETER_HOME = "C:\apache-jmeter-5.6.3"
.\scripts\generate-jmeter.ps1 -InputPath .\templates\api-input-template.csv -OutputDir .\out -Run
```

JMeter の場所は `--jmeter` / `-JMeterPath` でも指定できます。既存の JTL は追記されるため、`--run` と `-Run` は実行前に前回の `results.jtl`、`listener-results.jtl`、`jmeter.log`、`html-report` を削除します。

期待ステータスを書いた API だけ、応答コードの一致を検証します。書いていない API は、応答を受け取れた時点で JMeter 上は成功です。

## 制限

- メソッドは GET と POST です。
- `multipart/form-data` には対応していません。
- ヘッダ名は HTTP トークン文字だけです。ヘッダ値の改行は拒否します。
- Authorization などの秘密値は JMX に平文で入ります。利用手順の Excel では伏せます。
- JMX 内の JTL パスは出力ディレクトリからの相対パスです。GUI で開くときも、作業ディレクトリを出力ディレクトリにしてください。
