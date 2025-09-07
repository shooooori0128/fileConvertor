#!/usr/bin/env python3
"""
テスト用のサンプルExcelファイルを生成するスクリプト
"""

import pandas as pd
from pathlib import Path

def create_sample_excel():
    """サンプルExcelファイルを作成"""
    
    # サンプルデータの作成
    sample_data1 = {
        '商品名': ['りんご', 'バナナ', 'オレンジ', 'ぶどう', 'いちご'],
        '価格': [150, 120, 200, 300, 250],
        '在庫数': [50, 30, 20, 15, 25],
        '産地': ['青森', 'フィリピン', 'アメリカ', '山梨', '栃木']
    }
    
    sample_data2 = {
        '月': ['1月', '2月', '3月', '4月', '5月'],
        '売上': [1000000, 1200000, 1500000, 1300000, 1100000],
        '利益率': [0.15, 0.18, 0.20, 0.17, 0.16]
    }
    
    sample_data3 = {
        '名前': ['田中', '佐藤', '鈴木', '高橋'],
        '部署': ['営業', '開発', '営業', '人事'],
        '年齢': [28, 32, 25, 45],
        '経験年数': [3, 8, 2, 20]
    }
    
    # DataFrameを作成
    df1 = pd.DataFrame(sample_data1)
    df2 = pd.DataFrame(sample_data2)
    df3 = pd.DataFrame(sample_data3)
    
    # Excelファイルに書き出し
    output_file = Path('./input/sample_data.xlsx')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='商品マスタ', index=False)
        df2.to_excel(writer, sheet_name='月別売上', index=False)
        df3.to_excel(writer, sheet_name='従業員情報', index=False)
    
    print(f"サンプルExcelファイルを作成しました: {output_file}")
    print("シート:")
    print("  - 商品マスタ")
    print("  - 月別売上") 
    print("  - 従業員情報")

if __name__ == '__main__':
    create_sample_excel()
