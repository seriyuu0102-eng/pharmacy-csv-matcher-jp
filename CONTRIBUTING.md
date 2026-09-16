# Contributing

Thanks for helping improve Pharmacy CSV Matcher JP.

## Good contributions

Focused contributions are especially useful in these areas:

- CSV header-name compatibility
- matching and normalization edge cases
- clearer error messages
- Excel/VBA usability improvements
- documentation for pharmacy and healthcare workflows
- reproducible sample data that demonstrates a bug

## Before opening an issue

Please include:

1. what you expected to happen
2. what happened instead
3. the relevant column names or CSV structure
4. your Excel/Windows environment when it matters
5. a minimal reproducible example using dummy data only

## Privacy and confidential data

Do **not** upload or paste:

- patient information
- real prescription data
- real pharmacy inventory exports
- store-identifying internal files
- purchasing conditions or pricing
- wholesaler-confidential data
- credentials, tokens, or internal URLs

If a real CSV exposed a problem, recreate the same structure with fictional names and values before sharing it.

## Pull requests

Keep pull requests narrow and explain:

- the problem being solved
- the behavior before the change
- the behavior after the change
- how the change was tested

For VBA changes, please test the module by importing it into a macro-enabled Excel workbook and running the sample workflow with the dummy CSV files in `pharmacy_csv_matcher_sample/`.

## Matching philosophy

The project intentionally avoids relying only on JAN/product codes because spreadsheet and CSV handling can alter long numeric identifiers. Changes to the matching strategy should preserve transparency and make false matches easy to review.

## License

By contributing, you agree that your contribution will be licensed under the repository's MIT License.
