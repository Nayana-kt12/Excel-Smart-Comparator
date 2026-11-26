import unittest
from io import BytesIO
from unittest.mock import patch

from openpyxl import load_workbook
import pandas as pd

from bedrock_mapper import suggest_column_mapping
from comparator import (
    build_column_mapping,
    compare_workbooks,
    filter_comparison_columns,
    generate_excel_report,
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

    def test_employee_code_matches_rows_and_preserves_first_file_column_order(self):
        first = pd.DataFrame({
            "Emp Code": ["A1", "A2"],
            "Employee Name": ["Alex", "Blair"],
            "Department": ["Assembly", "Finishing"],
            "First Only": ["x", "y"],
        })
        second = pd.DataFrame({
            "BIO NO.": ["A2", "A1", "A3"],
            "Name": ["Blair", "Alex", "Casey"],
            "DEPT": ["Stores", "Assembly", "Paint"],
            "Second Only": [2, 1, 3],
        })

        result = compare_workbooks(first, second)

        self.assertEqual(int(result["summary"].iloc[0]["Matched rows"]), 2)
        self.assertEqual(
            result["summary"].iloc[0]["Row matching method"],
            "ID-based",
        )
        self.assertEqual(
            result["side_by_side"].columns.tolist(),
            [
                "Emp Code",
                "BIO NO.",
                "Employee Name",
                "Name",
                "Department",
                "DEPT",
            ],
        )
        self.assertNotIn(
            "First Only",
            result["side_by_side"].columns,
        )
        self.assertNotIn(
            "Second Only",
            result["side_by_side"].columns,
        )
        first_codes = result["side_by_side"]["Emp Code"].tolist()
        second_codes = result["side_by_side"]["BIO NO."].tolist()
        self.assertEqual(first_codes[:2], ["A1", "A2"])
        self.assertTrue(pd.isna(first_codes[2]))
        self.assertEqual(second_codes, ["A1", "A2", "A3"])
        self.assertEqual(
            result["side_by_side"]["Name"].tolist(),
            ["Alex", "Blair", "Casey"],
        )

    def test_employee_code_is_preferred_when_another_id_column_exists(self):
        first = pd.DataFrame({
            "Employee ID": ["ID-1"],
            "Emp Code": ["A1"],
            "Name": ["Alex"],
        })
        second = pd.DataFrame({
            "Employee ID": ["ID-1"],
            "BIO NO.": ["A2"],
            "Name": ["Alex"],
        })

        result = compare_workbooks(first, second)

        self.assertEqual(int(result["summary"].iloc[0]["Matched rows"]), 0)
        self.assertEqual(
            result["side_by_side_key_columns"],
            ["Emp Code", "BIO NO."],
        )

    def test_report_contains_comparison_sheets_and_mapping_details(self):
        first = pd.DataFrame({
            "Emp Code": ["A1"],
            "Name": ["Alex"],
            "First Only": ["x"],
        })
        second = pd.DataFrame({
            "BIO NO.": ["A1"],
            "Name": ["Alex"],
            "Second Only": ["y"],
        })
        result = compare_workbooks(first, second)
        report_bytes = generate_excel_report(result)

        with pd.ExcelFile(BytesIO(report_bytes)) as workbook:
            self.assertEqual(
                workbook.sheet_names,
                [
                    "Summary",
                    "Column Comparison",
                    "Row Comparison",
                    "Unmatched Rows",
                    "comparision",
                ],
            )
            column_report = pd.read_excel(
                workbook,
                sheet_name="Column Comparison",
            )
            self.assertIn("First Only", column_report["First workbook column"].tolist())
            self.assertIn("Second Only", column_report["Second workbook column"].tolist())
            self.assertIn("Matched", column_report["Column status"].tolist())
            self.assertIn("No match found", column_report["Column status"].tolist())
            summary_report = pd.read_excel(
                workbook,
                sheet_name="Summary",
            )
            self.assertEqual(
                summary_report.iloc[0]["First workbook columns"],
                3,
            )
            self.assertEqual(
                summary_report.iloc[0]["Second workbook columns"],
                3,
            )

    def test_unmatched_rows_use_source_colors_without_source_tabs(self):
        first = pd.DataFrame({
            "Emp Code": ["A1", "A2"],
            "Employee Name": ["Alex", "Blair"],
            "Department": ["Assembly", "Finishing"],
            "Salary": [1200, 1500],
        })
        second = pd.DataFrame({
            "BIO NO.": ["A2", "A3"],
            "Name": ["Blair", "Casey"],
            "DEPT": ["Stores", "Paint"],
            "Salary": [1600, 1800],
        })
        result = compare_workbooks(first, second)
        report_bytes = generate_excel_report(result)

        workbook = load_workbook(BytesIO(report_bytes))
        worksheet = workbook["comparision"]
        self.assertEqual(
            workbook.sheetnames,
            [
                "Summary",
                "Column Comparison",
                "Row Comparison",
                "Unmatched Rows",
                "comparision",
            ],
        )
        for sheet_name in ("Summary", "comparision"):
            header_cell = workbook[sheet_name]["A1"]
            self.assertTrue(header_cell.font.bold)
            self.assertEqual(header_cell.alignment.horizontal, "center")
        self.assertNotIn(
            "Unmatched",
            [cell.value for cell in worksheet[1]],
        )
        self.assertEqual(worksheet["A2"].fill.fgColor.rgb[-6:], "C00060")
        self.assertEqual(worksheet["C2"].fill.fgColor.rgb[-6:], "F4CCCC")
        self.assertEqual(worksheet["E2"].fill.fgColor.rgb[-6:], "F4CCCC")
        self.assertEqual(worksheet["G2"].fill.fgColor.rgb[-6:], "FCE4D6")
        self.assertEqual(worksheet["B4"].fill.fgColor.rgb[-6:], "C00060")
        self.assertEqual(worksheet["D4"].fill.fgColor.rgb[-6:], "D9D2E9")
        self.assertEqual(worksheet["F4"].fill.fgColor.rgb[-6:], "D9D2E9")
        self.assertEqual(worksheet["H4"].fill.fgColor.rgb[-6:], "DDEBF7")
        self.assertEqual(worksheet["E3"].fill.fgColor.rgb[-6:], "FFD700")
        self.assertEqual(worksheet["F3"].fill.fgColor.rgb[-6:], "FFD700")
        self.assertEqual(worksheet["G3"].fill.fgColor.rgb[-6:], "FFD700")
        self.assertEqual(worksheet["H3"].fill.fgColor.rgb[-6:], "FFD700")

    def test_blank_mapped_value_does_not_get_difference_highlighting(self):
        first = pd.DataFrame({
            "Emp Code": ["A1"],
            "Department": ["House Keeping"],
        })
        second = pd.DataFrame({
            "BIO NO.": ["A1"],
            "DEPT": [" "],
        })
        report_bytes = generate_excel_report(
            compare_workbooks(first, second)
        )

        worksheet = load_workbook(BytesIO(report_bytes))["comparision"]
        self.assertIsNone(worksheet["C2"].fill.fill_type)
        self.assertIsNone(worksheet["D2"].fill.fill_type)

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
            "BASIC": [100],
            "Basic": [90],
        })

        mapping = build_column_mapping(first, second)
        result = compare_workbooks(first, second, mapping=mapping, threshold=0.5)

        self.assertEqual(
            mapping["Basic"],
            ["BASIC", "Basic"],
        )
        self.assertEqual(
            result["columns"]["Second workbook column"].tolist(),
            ["BASIC", "Basic"],
        )
        self.assertEqual(
            result["rows"]["Difference: Basic | BASIC"].tolist(),
            ["Matched"],
        )
        self.assertEqual(
            result["rows"]["Difference: Basic | Basic"].tolist(),
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

    def test_fields_with_special_names_remain_in_comparison_results(self):
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

        self.assertEqual(
            list(filter_comparison_columns(first).columns),
            ["Employee Name", "DOJ", "ESIC NO", "UAN NO"],
        )
        self.assertEqual(
            list(filter_comparison_columns(second).columns),
            ["Name", "DOJ", "ESIC NO", "UAN NO"],
        )
        self.assertIn(
            "DOJ",
            result["columns"]["First workbook column"].tolist(),
        )
        self.assertIn(
            "ESIC NO",
            result["columns"]["First workbook column"].tolist(),
        )
        self.assertIn(
            "UAN NO",
            result["columns"]["First workbook column"].tolist(),
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
                "Basic",
                "HRA",
                "Basic",
                "HRA",
            ],
        )
        self.assertEqual(columns, list(frame.columns))

    def test_preserves_special_column_names_when_reading_sheets(self):
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

        self.assertEqual(
            list(frame.columns),
            ["Name", "DOJ", "ESIC NO", "UAN NO"],
        )
        self.assertEqual(columns, ["Name", "DOJ", "ESIC NO", "UAN NO"])


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
