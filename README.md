# Pharmacy CSV Matcher JP

Open-source Excel/VBA CSV matching tool for Japanese community pharmacies.

## Overview

Pharmacy CSV Matcher JP is an Excel/VBA tool designed to help Japanese community pharmacies compare multiple CSV files, such as adopted medicine lists, inventory data, and target medicine lists.

The tool is intended to reduce manual checking work and prevent matching errors caused by unstable product codes, scientific notation, lost leading zeros, and CSV formatting issues.

## Features

- Match adopted item lists and inventory CSV files
- Compare against target medicine lists
- Export matched results to an Excel worksheet
- Avoid failures caused by scientific notation in product codes
- Use product name, specification, and manufacturer as stable matching keys
- Provide sample CSV files for testing

## Included Sample Files

The sample files are included in the `pharmacy_csv_matcher_sample` folder.

- `sample_adopted_items.csv`
- `sample_inventory.csv`
- `sample_target_items.csv`
- `PharmacyCsvMatcher.bas`

All sample files contain dummy data only.

## Intended Users

- Community pharmacists
- Pharmacy managers
- Medical inventory administrators
- Healthcare professionals who use Excel-based workflows

## Use Cases

This tool may be useful when a pharmacy needs to:

- Check whether stocked items are included in an adopted item list
- Compare inventory data against a target medicine list
- Reduce manual lookup work
- Prepare internal reports based on CSV exports
- Avoid errors caused by product code formatting problems

## How to Use

1. Open Excel.
2. Create or open a macro-enabled workbook.
3. Open the VBA editor.
4. Import `PharmacyCsvMatcher.bas` from the `pharmacy_csv_matcher_sample` folder.
5. Run `RunPharmacyCsvMatcher`.
6. Select the adopted item CSV, inventory CSV, and target item CSV.
7. Review the output worksheet.

## Important Notes

This project does not include real pharmacy data, patient data, or confidential business data.

The included CSV files are sample files for demonstration and testing purposes only.

## Roadmap

- Improve CSV format auto-detection
- Add better error messages
- Add a simple user interface
- Support more flexible column name mapping

- 
- Add export templates for reports
- Improve documentation for non-technical users
- ## Screenshots

### VBA module import
![VBA import](images/vba-import-sample.png)

### CSV file selection
![CSV selection](images/csv-selection-sample.png)

### Output worksheet example
![Excel output](images/excel-output-sample.png)

## Development Status

This project is in an early development stage.
Planned improvements are tracked in GitHub Issues, including better CSV column detection, Japanese documentation, and Excel workflow screenshots.

## License

MIT License

---

## 日本語説明

Pharmacy CSV Matcher JP は、日本の薬局業務で発生しやすいCSV突合作業を支援するExcel/VBAツールです。

採用品CSV、在庫CSV、対象品目リストCSVを照合し、条件に一致する品目をExcelシートに出力します。

薬局現場では、商品コードやJANコードがCSV上で指数表記になったり、先頭ゼロが消えたり、Excel上で丸められたりすることがあります。そのため、このツールでは商品コードだけに頼らず、品名・規格容量・メーカーを組み合わせたキーで照合する設計にしています。

## 想定している利用者

- 薬局薬剤師
- 薬局管理者
- 在庫管理担当者
- Excelで薬局内データを管理している人

## 注意事項

このリポジトリには、実店舗の在庫データ、患者情報、卸由来の非公開データは含まれていません。

含まれているCSVは、公開用に作成したダミーデータです。
