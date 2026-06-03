# Pharmacy CSV Matcher サンプル一式

このフォルダは、薬局業務で発生しやすい「採用品CSV」「在庫CSV」「対象品目CSV」の突合を公開OSS化するための最小サンプルです。

## ファイル

- `sample_adopted_items.csv`
  - 採用品CSVのサンプル
- `sample_inventory.csv`
  - 在庫CSVのサンプル
- `sample_target_items.csv`
  - 選定療養など、別リスト側のサンプル
- `PharmacyCsvMatcher.bas`
  - Excel VBA用の標準モジュール

## 想定する突合条件

商品コード/JANは、CSVやExcelで `4.987E+12` のような指数表記になったり、先頭ゼロが消えたりする可能性があります。  
そのため、このサンプルでは商品コードを主キーにせず、次の3項目を突合キーにしています。

```text
品名 + 規格容量 + メーカー
```

## 使い方

1. Excelで新しい `.xlsm` ブックを作成する
2. VBAエディタを開く
3. `PharmacyCsvMatcher.bas` をインポートする
4. `RunPharmacyCsvMatcher` を実行する
5. 順番に以下のCSVを選択する
   - 採用品CSV
   - 在庫CSV
   - 対象品目CSV
6. Excel内に `突合結果` と `突合ログ` シートが作成される

## 出力される条件

以下すべてを満たす行だけを `突合結果` に出力します。

- 在庫CSVで在庫数が1以上
- 採用品CSVと一致
- 対象品目CSVと一致

## 公開時の注意

実店舗名、患者情報、卸由来の非公開データ、実在の仕入れ条件などは入れないでください。  
GitHubに公開する場合は、必ず架空データ・ダミーデータのみを使ってください。
