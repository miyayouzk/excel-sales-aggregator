# Excel Sales Aggregator

複数の売上Excelファイルを自動で結合し、

担当者別・商品別・日別・クロス集計を自動生成するPythonツールです。

## Features

\- 複数Excelの自動結合

\- 担当者別売上集計

\- 商品別売上集計

\- 日別売上集計

\- 担当者 × 日付 クロス集計

\- 1つのExcelに複数シートで出力

## How to Use

1\. input フォルダに売上Excelを入れる

2\. Excelまとめ.py を実行

3\. output フォルダに集計Excelが生成される

## Target Users

- 飲食店・バーなどの小規模店舗オーナー
- Excelで売上管理をしている方
- 手作業の集計を自動化したい方

## Input Format

以下の列を持つExcelファイルを想定しています。

- 日付
- 担当者
- 商品名
- 数量
- 単価
- 合計金額

## Tech Stack

- Python
- pandas
- openpyxl

