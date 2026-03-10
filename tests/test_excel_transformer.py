import tempfile
import unittest
from pathlib import Path
from importlib.util import find_spec

HAS_EXCEL_DEPS = find_spec("pandas") is not None and find_spec("openpyxl") is not None

if not HAS_EXCEL_DEPS:
    pd = None
    ExcelTransformationService = None
else:
    import pandas as pd
    from excel_transformer import ExcelTransformationService


class ExcelTransformationServiceTests(unittest.TestCase):
    @unittest.skipIf(not HAS_EXCEL_DEPS, "pandas/openpyxl are not installed in this environment")
    def test_transform_workbook_applies_rules_and_template_order(self):
        data_df = pd.DataFrame(
            [
                {"FirstName": "jane", "LastName": "doe", "Country": "US", "Score": 93, "RunDate": "2025-01-15"},
                {"FirstName": "john", "LastName": "smith", "Country": "CA", "Score": 72, "RunDate": "2025-01-16"},
            ]
        )
        criteria_df = pd.DataFrame(
            [
                {"rule_order": 1, "target_column": "Full Name", "source_column": "FirstName,LastName", "operation": "CONCAT", "operation_arg": " ", "required": True},
                {"rule_order": 2, "target_column": "Country", "source_column": "Country", "operation": "UPPER"},
                {"rule_order": 3, "target_column": "Tier", "operation": "STATIC", "operation_arg": "A", "condition": "Score >= 90"},
                {"rule_order": 4, "target_column": "Tier", "operation": "STATIC", "operation_arg": "B", "condition": "Score < 90"},
                {"rule_order": 5, "target_column": "Run Date", "source_column": "RunDate", "operation": "DATE_FORMAT", "operation_arg": "%Y/%m/%d"},
            ]
        )
        result_template_df = pd.DataFrame(columns=["Full Name", "Tier", "Country", "Run Date"])

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "input.xlsx"
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                data_df.to_excel(writer, sheet_name="Data", index=False)
                criteria_df.to_excel(writer, sheet_name="Criteria", index=False)
                result_template_df.to_excel(writer, sheet_name="Result", index=False)

            service = ExcelTransformationService()
            result_df = service.transform_workbook(path, path)

            self.assertEqual(list(result_df.columns), ["Full Name", "Tier", "Country", "Run Date"])
            self.assertEqual(result_df.loc[0, "Full Name"], "jane doe")
            self.assertEqual(result_df.loc[0, "Tier"], "A")
            self.assertEqual(result_df.loc[1, "Tier"], "B")
            self.assertEqual(result_df.loc[0, "Run Date"], "2025/01/15")


if __name__ == "__main__":
    unittest.main()
