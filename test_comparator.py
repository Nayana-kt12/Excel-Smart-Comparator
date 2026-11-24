import unittest
from io import BytesIO
from unittest.mock import patch

import pandas as pd

from bedrock_mapper import suggest_column_mapping
from comparator import (
    build_column_mapping,
    compare_workbooks,
    filter_comparison_columns,
    values_equal,
)
from excel_reader import read_sheet_columns, read_sheet_with_detected_header


class ColumnMappingTests(unittest.TestCase):
    def test_department_matches_dept(self):
        first = pd.DataFrame(columns=["Department"])
        second = pd.DataFrame(columns=["DEPT"])

        self.assertEqual(
            build_column_mapping(first, second),
            {"Department": "DEPT"},
        )

    def test_basic_matches_bs(self):
        first = pd.DataFrame(columns=["Basic"])
        second = pd.DataFrame(columns=["BS"])

        self.assertEqual(build_column_mapping(first, second), {"Basic": "BS"})


    def test_excel_timestamp_matches_iso_date_string(self):
        self.assertTrue(
            values_equal(
                pd.Timestamp("2024-01-15"),
                "2024-01-15",
            )
        )

 

    def test_employee_code_matches_bio_number(self):
        first = pd.DataFrame(columns=["Employee Code"])
        second = pd.DataFrame(columns=["BIO NO."])

        self.assertEqual(
            build_column_mapping(first, second),
            {"Employee Code": "BIO NO."},
        )

    def test_employee_name_matches_name_not_colliding_initialism(self):
        first = pd.DataFrame(columns=["Employee Name"])
        second = pd.DataFrame(columns=["ESIC NO", "Name"])

        self.assertEqual(
            build_column_mapping(first, second),
            {"Employee Name": "Name"},
        )

    def test_ambiguous_alias_is_not_guessed(self):
        first = pd.DataFrame(columns=["Basic"])
        second = pd.DataFrame(columns=["BS", "Basic Salary"])

        self.assertEqual(build_column_mapping(first, second), {"Basic": None})

    def test_repeated_grouped_headers_map_to_each_matching_column(self):
        first = pd.DataFrame({"Basic": [100]})
        second = pd.DataFrame({
            "Master Wages | BASIC": [100],
            "Earning Wages | BASIC": [90],
        })

        mapping = build_column_mapping(first, second)
        result = compare_workbooks(first, second, mapping=mapping, threshold=0.5)

        self.assertEqual(
            mapping["Basic"],
            ["Master Wages | BASIC", "Earning Wages | BASIC"],
        )
        self.assertEqual(
            result["columns"]["Second workbook column"].tolist(),
            ["Master Wages | BASIC", "Earning Wages | BASIC"],
        )
        self.assertEqual(
            result["rows"]["Difference: Basic | Master Wages | BASIC"].tolist(),
            ["Matched"],
        )
        self.assertEqual(
            result["rows"]["Difference: Basic | Earning Wages | BASIC"].tolist(),
            [-10.0],
        )

    def test_unique_names_allow_department_values_to_be_compared(self):
        first = pd.DataFrame({
            "Employee Name": ["Alex", "Blair"],
            "Department": ["Assembly", "Finishing"],
        })
        second = pd.DataFrame({
            "Name": ["Alex", "Blair"],
            "DEPT": ["Assembly", "Stores"],
        })

        result = compare_workbooks(first, second, threshold=1.0)

        self.assertEqual(int(result["summary"].iloc[0]["Matched rows"]), 2)
        self.assertEqual(
            result["rows"]["Difference: Department | DEPT"].tolist(),
            ["Matched", "No match found"],
        )

    def test_duplicate_names_are_not_used_as_row_keys(self):
        first = pd.DataFrame({
            "Employee Name": ["Alex", "Alex"],
            "Department": ["Assembly", "Finishing"],
        })
        second = pd.DataFrame({
            "Name": ["Alex"],
            "DEPT": ["Stores"],
        })

        result = compare_workbooks(first, second, threshold=1.0)

        self.assertEqual(int(result["summary"].iloc[0]["Matched rows"]), 0)

    def test_excluded_fields_are_omitted_from_comparison_results(self):
        first = pd.DataFrame({
            "Employee Name": ["Alex"],
            "DOJ": ["2023-02-14"],
            "ESIC NO": ["3314851913"],
            "UAN NO": ["101601952623"],
        })
        second = pd.DataFrame({
            "Name": ["Alex"],
            "DOJ": ["2023-02-14"],
            "ESIC NO": ["3314851913"],
            "UAN NO": ["101601952623"],
        })

        result = compare_workbooks(first, second)

        self.assertEqual(list(filter_comparison_columns(first).columns), ["Employee Name"])
        self.assertEqual(list(filter_comparison_columns(second).columns), ["Name"])
        for sheet in ("columns", "rows", "unmatched"):
            self.assertFalse(
                any(
                    any(name in str(column).casefold() for name in ("doj", "esic no", "uan no"))
                    for column in result[sheet].columns
                ),
                sheet,
            )


class ExcelReaderTests(unittest.TestCase):
    def test_detects_text_headers_above_data_rows(self):
        workbook_bytes = BytesIO()
        with pd.ExcelWriter(workbook_bytes, engine="openpyxl") as writer:
            pd.DataFrame([
                ["Wage report", None, None, None],
                ["Master wage", None, "Earning wage", None],
                ["SR. No.", "BIO NO.", "Name", "DEPT"],
                [1, 1001, "Alex", "Assembly"],
            ]).to_excel(writer, sheet_name="Wages", header=False, index=False)
        workbook_bytes.seek(0)

        with pd.ExcelFile(workbook_bytes) as workbook:
            frame = read_sheet_with_detected_header(workbook, "Wages")

        self.assertEqual(list(frame.columns), ["SR. No.", "BIO NO.", "Name", "DEPT"])
        self.assertEqual(frame.iloc[0].tolist(), [1, 1001, "Alex", "Assembly"])
        self.assertEqual(frame.attrs["detected_header_row"], 3)

    def test_preserves_repeated_headers_under_their_main_headers(self):
        workbook_bytes = BytesIO()
        with pd.ExcelWriter(workbook_bytes, engine="openpyxl") as writer:
            pd.DataFrame([
                ["Report", None, None, None, None],
                [None, "Master Wages", None, "Earning Wages", None],
                ["Name", "Basic", "HRA", "Basic", "HRA"],
                ["Alex", 100, 20, 90, 18],
            ]).to_excel(writer, sheet_name="Wages", header=False, index=False)
            worksheet = writer.book["Wages"]
            worksheet.merge_cells("B2:C2")
            worksheet.merge_cells("D2:E2")
        workbook_bytes.seek(0)

        with pd.ExcelFile(
            workbook_bytes,
            engine_kwargs={"read_only": False},
        ) as workbook:
            frame = read_sheet_with_detected_header(workbook, "Wages")
            columns = read_sheet_columns(workbook, "Wages")

        self.assertEqual(
            list(frame.columns),
            [
                "Name",
                "Master Wages | Basic",
                "Master Wages | HRA",
                "Earning Wages | Basic",
                "Earning Wages | HRA",
            ],
        )
        self.assertEqual(columns, list(frame.columns))

    def test_excludes_requested_columns_from_read_sheets(self):
        workbook_bytes = BytesIO()
        with pd.ExcelWriter(workbook_bytes, engine="openpyxl") as writer:
            pd.DataFrame([{
                "Name": "Alex",
                "DOJ": "2023-02-14",
                "ESIC NO": "3314851913",
                "UAN NO": "101601952623",
            }]).to_excel(writer, sheet_name="Wages", index=False)
        workbook_bytes.seek(0)

        with pd.ExcelFile(workbook_bytes) as workbook:
            frame = read_sheet_with_detected_header(workbook, "Wages")
            columns = read_sheet_columns(workbook, "Wages")

        self.assertEqual(list(frame.columns), ["Name"])
        self.assertEqual(columns, ["Name"])


class BedrockMapperTests(unittest.TestCase):
    @patch("bedrock_mapper.boto3.client")
    def test_converse_request_omits_deprecated_temperature(self, create_client):
        client = create_client.return_value
        client.converse.return_value = {
            "output": {
                "message": {
                    "content": [{"text": '{"Department":"DEPT"}'}]
                }
            }
        }

        result = suggest_column_mapping(
            ["Department"],
            ["DEPT"],
            "global.anthropic.claude-sonnet-5-5",
            "us-east-1",
        )

        self.assertEqual(result, {"Department": "DEPT"})
        self.assertEqual(
            client.converse.call_args.kwargs["inferenceConfig"],
            {"maxTokens": 3000},
        )

    

        result = suggest_column_mapping(
            ["Name"],
            "global.anthropic.claude-sonnet-5-5",
            "us-east-1",
        )

        self.assertEqual(result, {})

    @patch("bedrock_mapper.boto3.client")
    def test_truncated_response_has_clear_error(self, create_client):
        create_client.return_value.converse.return_value = {
            "stopReason": "max_tokens",
            "output": {"message": {"content": []}},
        }

        with self.assertRaisesRegex(ValueError, "response was truncated"):
            suggest_column_mapping(["Department"], ["DEPT"])


if __name__ == "__main__":
    unittest.main()
