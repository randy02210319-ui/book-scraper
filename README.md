# 書籍一覧スクレイピングツール（Python）

Webサイトの一覧ページを自動でたどり、**商品名・価格・評価・在庫・URL** を集めて **CSV と Excel** に書き出すツールです。

「競合サイトの価格を毎回手作業でコピーしている」「掲載情報を一覧表にまとめたい」といった作業を自動化できます。

> 対象サイトは、スクレイピング練習用に公開されている [books.toscrape.com](https://books.toscrape.com/) です。

## できること

| 機能 | 内容 |
|---|---|
| 複数ページの自動巡回 | 「次へ」のリンクをたどり、指定したページ数（または最後のページ）まで自動で取得 |
| CSV 出力 | Excel で開いても日本語が文字化けしない形式（UTF-8 BOM付き）で保存 |
| Excel 出力 | 見出しの色付け・固定、絞り込み（フィルター）、価格の通貨表示、列幅の調整まで済ませた状態で保存 |
| エラーに強い | 一時的な通信エラーのときは、間隔をあけて自動で最大3回やり直す |
| サーバーへの配慮 | ページごとに待ち時間を入れて、相手のサイトに負担をかけない |

## 実行結果

全50ページ・1000冊を、待ち時間0.5秒の設定で約45秒かけて取得しました。

```
$ python scraper.py --pages 0 --delay 0.5
23:56:01 1ページ目を取得中: https://books.toscrape.com/catalogue/page-1.html
23:56:02 2ページ目を取得中: https://books.toscrape.com/catalogue/page-2.html
...
23:56:43 50ページ目を取得中: https://books.toscrape.com/catalogue/page-50.html
23:56:44 1000冊を保存しました: output/books_20261005_235644.csv / output/books_20261005_235644.xlsx
```

出力サンプル（先頭20件）: [docs/sample.csv](docs/sample.csv)

| タイトル | 価格（£） | 評価（5段階） | 在庫 | URL |
|---|---|---|---|---|
| A Light in the Attic | 51.77 | 3 | あり | https://books.toscrape.com/catalogue/a-light-in-the-attic_1000/index.html |
| Tipping the Velvet | 53.74 | 1 | あり | https://books.toscrape.com/catalogue/tipping-the-velvet_999/index.html |
| Soumission | 50.1 | 1 | あり | https://books.toscrape.com/catalogue/soumission_998/index.html |

## 使用技術

- Python 3.9 以上
- requests（ページの取得、自動再試行）
- Beautiful Soup 4（HTML の解析）
- openpyxl（Excel ファイルの作成）

## セットアップ

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 使い方

```bash
python scraper.py              # 1ページ目（20冊）だけ取得
python scraper.py --pages 5    # 5ページ（100冊）取得
python scraper.py --pages 0    # 全ページ取得
```

| オプション | 説明 | 初期値 |
|---|---|---|
| `--pages` | 取得するページ数（0 で全ページ） | `1` |
| `--delay` | ページを取得する間隔（秒） | `1.0` |
| `--out-dir` | 出力先フォルダ | `output` |

ファイル名には実行日時が付く（例: `books_20261005_235644.csv`）ので、何度実行しても前の結果が上書きされません。

## 工夫した点

- **省略されないタイトルを取得**：一覧ページでは長いタイトルが「A Light in the ...」と省略表示されるため、HTML の `title` 属性から正式なタイトルを取り出している
- **文字化け対策**：サイトの文字コードが誤って判定され「£」が化ける問題を、文字コードを明示して解消。CSV も Excel で開いて文字化けしない形式にした
- **そのまま使える Excel**：受け取った方がすぐに並べ替え・絞り込みできるよう、書式を整えた状態で出力
- **相手のサイトへの配慮**：アクセスの間隔をあけ、ツール名を名乗る User-Agent を設定

## 注意

スクレイピングを行う際は、対象サイトの利用規約と `robots.txt` を確認し、許可されている範囲で利用してください。
