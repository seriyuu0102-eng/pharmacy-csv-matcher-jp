# Changelog

All notable changes to Pharmacy CSV Matcher JP will be documented here.

## Unreleased

### Added

- common Japanese and English CSV header aliases for product name, package size, manufacturer, inventory quantity, store name, product code, and optional output fields
- contribution guidelines with privacy requirements for pharmacy-related examples
- reproducible Japanese-alias and English-alias CSV compatibility fixtures
- expected-match files and a manual integration-test matrix under `tests/`

### Changed

- Excel display settings are restored when processing fails
- README now documents project scope, supported aliases, screenshots, privacy expectations, maintenance focus, and compatibility testing

### Fixed

- README screenshot paths now point to the actual files in the `image/` directory

## Initial public version

- Excel/VBA matching module
- dummy adopted-item, inventory, and target-item CSV files
- matching based on normalized product name, package size, and manufacturer
- result and processing-log worksheets
