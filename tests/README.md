# Compatibility test fixtures

This directory contains fictional CSV fixtures for manually verifying `PharmacyCsvMatcher.bas` in Excel.

No patient data, real pharmacy data, wholesaler data, purchasing conditions, or real store identifiers are included.

## Test matrix

| Fixture | Purpose | Expected output rows |
| --- | --- | ---: |
| `pharmacy_csv_matcher_sample/` | canonical Japanese headers and matching edge cases | 4 |
| `tests/fixtures/ja_alias/` | common Japanese header aliases | 2 |
| `tests/fixtures/en_alias/` | common English header aliases | 2 |

## How to run a fixture

1. Create or open a macro-enabled Excel workbook (`.xlsm`).
2. Import `pharmacy_csv_matcher_sample/PharmacyCsvMatcher.bas`.
3. Run `RunPharmacyCsvMatcher`.
4. When prompted, select the fixture's `adopted.csv`, `inventory.csv`, and `target.csv` in that order.
5. Confirm the `突合結果` worksheet contains the expected number of rows shown above.
6. Compare the matched products with `expected_matches.csv` in the fixture directory.
7. Confirm `突合ログ` reports the same output count.

## Japanese alias fixture

The `ja_alias` fixture intentionally uses alternate column names such as:

- `商品名` / `医薬品名`
- `規格` / `包装規格`
- `製造販売元` / `メーカー名`
- `在庫数量`
- `薬局名`
- `区分` / `先発名` / `メモ`

It also checks the existing manufacturer normalization by matching `架空ファーマ株式会社` against `架空ファーマ(株)`.

Expected matches: `テストA錠5mg` and `テストB錠10mg` only.

## English alias fixture

The `en_alias` fixture intentionally mixes supported English aliases such as:

- `product name`, `item name`, `drug name`
- `package size`, `specification`, `spec`
- `manufacturer`, `maker`
- `stock quantity`
- `store name`
- `target category`, `brand name`, `notes`

Expected matches: `Sample Alpha 5mg` and `Sample Beta 10mg` only.

## Encoding

The committed fixture CSV files are UTF-8 text with a BOM so that Japanese Excel can identify the encoding reliably in common environments.

`PharmacyCsvMatcher.bas` currently opens CSV files through Excel using `Workbooks.Open(..., Local:=True)`. CP932 / Shift-JIS behavior therefore depends on the local Excel and Windows environment and is not guaranteed by these fixtures. If CP932 compatibility becomes a requirement, add an explicit Windows Excel smoke-test case and document the supported environment.

## What these tests do not prove

These are reproducible manual integration fixtures, not an automated execution of VBA in CI. A passing fixture means the workflow behaved as expected in the Excel environment where it was run. Changes to header mapping, CSV loading, text normalization, or matching rules should be checked against all three fixture groups before release.
