# fileConvertor

🚨このリポジトリはAI Agentに書かせたコードです🚨

ExcelファイルをMarkdown形式に変換するCLIアプリケーションです。

## ディレクトリ構成

```
convertor/
├── input/          # Excelファイル(.xlsx)を配置
├── output/         # 変換されたMarkdownファイルの出力先
├── excel_to_markdown.py  # メインコンバーター
├── create_sample.py      # サンプルファイル生成スクリプト
└── requirements.txt      # 依存関係
```

## 機能

- Excelファイル（.xlsx）をMarkdownテーブル形式に変換
- 複数シートの対応
- CLI形式での実行
- カスタマイズ可能な出力フォーマット
- 入力ファイルは`input/`ディレクトリから読み込み
- 出力ファイルは自動的に`output/`ディレクトリに保存

## インストール

```bash
# 依存関係のインストール
pip install -r requirements.txt
```

## 使用方法

```bash
# 基本的な使用方法（inputディレクトリからファイルを読み込み、outputに保存）
python3 excel_to_markdown.py input/sample_data.xlsx

# 出力ファイル名を指定（outputディレクトリ内に保存）
python3 excel_to_markdown.py input/sample_data.xlsx -o custom_name.md

# 特定のシートのみ変換
python3 excel_to_markdown.py input/sample_data.xlsx -s "Sheet1"

# シート一覧を表示
python3 excel_to_markdown.py input/sample_data.xlsx --list-sheets

# サンプルファイルの作成
python3 create_sample.py
```

## 依存関係

- Python 3.7+
- openpyxl
- pandas
- click
- tabulate

## ライセンス

MIT License
