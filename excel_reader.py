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
    if header_row == 0:
        return frame

    workbook = getattr(excel_file, "book", None)
    sheet_names = getattr(workbook, "sheetnames", [])
    if sheet_name not in sheet_names:
        return frame

    worksheet = workbook[sheet_name]
    merged_cells = getattr(worksheet, "merged_cells", None)
    if merged_cells is None:
        return frame

    groups_by_column = {}
    for merged_range in merged_cells.ranges:
        if merged_range.min_row != header_row or merged_range.max_col == merged_range.min_col:
            continue
        parent = worksheet.cell(
            row=merged_range.min_row,
            column=merged_range.min_col,
        ).value
        if isinstance(parent, str) and parent.strip():
            for column_index in range(
                merged_range.min_col, merged_range.max_col + 1
            ):
                groups_by_column[column_index - 1] = parent.strip()

    if len(set(groups_by_column.values())) < 2:
        return frame

    qualified_columns = []
    for index, column in enumerate(frame.columns):
        if index not in groups_by_column:
            qualified_columns.append(column)
            continue

        child = worksheet.cell(
            row=header_row + 1,
            column=index + 1,
        ).value
        child = str(child).strip() if child is not None else str(column)
        qualified_columns.append(f"{groups_by_column[index]} | {child}")

    frame = frame.copy()
    frame.columns = qualified_columns
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
