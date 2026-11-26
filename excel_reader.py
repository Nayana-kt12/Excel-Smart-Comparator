import pandas as pd

from comparator import filter_comparison_columns


def _detect_header_row(excel_file, sheet_name, preview_rows):
    preview = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=None,
        nrows=preview_rows,
    )
    if preview.empty:
        return None
    if not preview.notna().any().any():
        return None

    def header_score(row):
        labels = [
            value.strip()
            for value in row
            if isinstance(value, str) and value.strip()
        ]
        return (
            len({label.casefold() for label in labels}),
            len(labels),
            int(row.notna().sum()),
            -int(row.name),
        )

    return int(max(
        preview.index,
        key=lambda index: header_score(preview.loc[index])
    ))


def _apply_grouped_header_names(excel_file, sheet_name, header_row, frame):
    workbook = getattr(excel_file, "book", None)
    sheet_names = getattr(workbook, "sheetnames", [])
    if sheet_name not in sheet_names:
        return frame

    worksheet = workbook[sheet_name]
    worksheet_column_count = worksheet.max_column
    if len(frame.columns) > worksheet_column_count:
        frame = frame.iloc[:, :worksheet_column_count].copy()
    uploaded_names = [
        worksheet.cell(row=header_row + 1, column=index + 1).value
        for index in range(len(frame.columns))
    ]
    frame = frame.copy()
    frame.columns = [
        str(name).strip() if name is not None else str(column)
        for name, column in zip(uploaded_names, frame.columns)
    ]
    return frame


def read_sheet_columns(excel_file, sheet_name, preview_rows=20):
    header_row = _detect_header_row(excel_file, sheet_name, preview_rows)
    if header_row is None:
        return []

    frame = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=header_row,
        nrows=0,
    )
    frame = _apply_grouped_header_names(
        excel_file, sheet_name, header_row, frame
    )
    return list(filter_comparison_columns(frame).columns)


def read_sheet_with_detected_header(excel_file, sheet_name, preview_rows=20):
    header_row = _detect_header_row(excel_file, sheet_name, preview_rows)
    if header_row is None:
        return pd.DataFrame()

    frame = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=header_row,
    )
    frame = _apply_grouped_header_names(
        excel_file, sheet_name, header_row, frame
    )
    frame = filter_comparison_columns(frame)
    frame.attrs["detected_header_row"] = header_row + 1
    return frame
