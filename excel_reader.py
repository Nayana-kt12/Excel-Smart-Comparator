import pandas as pd

from comparator import (
    column_alias_group,
    column_field_name,
    filter_comparison_columns,
    norm_name,
)


def _group_header_name(worksheet, header_row, column_index):
    if header_row < 1:
        return None

    for merged_range in worksheet.merged_cells.ranges:
        if (
            merged_range.min_row == header_row
            and merged_range.min_col <= column_index <= merged_range.max_col
        ):
            value = worksheet.cell(
                row=merged_range.min_row,
                column=merged_range.min_col,
            ).value
            if value is not None and str(value).strip():
                return str(value).strip()
    return None


def _header_equivalence_key(name):
    field_name = norm_name(column_field_name(name))
    if field_name.startswith("pf"):
        return "statutory:pf"
    if field_name.startswith("esic"):
        return "statutory:esic"

    alias_group = column_alias_group(name)
    if alias_group is not None:
        return f"alias:{alias_group}"
    return f"name:{field_name}"


def _unique_column_names(columns, worksheet, header_row):
    names = [str(column) for column in columns]
    counts = {}
    for name in names:
        normalized = _header_equivalence_key(name)
        counts[normalized] = counts.get(normalized, 0) + 1

    for index, name in enumerate(names):
        normalized = _header_equivalence_key(name)
        if counts[normalized] > 1:
            group_name = _group_header_name(
                worksheet,
                header_row,
                index + 1,
            )
            if group_name:
                names[index] = f"{group_name} | {name}"

    used = set()
    unique_names = []
    for name in names:
        normalized = name.casefold()
        candidate = name
        if normalized in used:
            occurrence = 2
            candidate = f"{name} [duplicate {occurrence}]"
            while candidate.casefold() in used:
                occurrence += 1
                candidate = f"{name} [duplicate {occurrence}]"
        used.add(candidate.casefold())
        unique_names.append(candidate)
    return unique_names


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
    frame.columns = _unique_column_names(
        frame.columns,
        worksheet,
        header_row,
    )
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
