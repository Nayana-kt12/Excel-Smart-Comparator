import re
import math
import datetime
import pandas as pd
from difflib import SequenceMatcher
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment

NO_MATCH = "No match found"
EXCLUDED_COMPARISON_COLUMNS = frozenset()

COLUMN_ALIAS_GROUPS = (
    ("basic", "basic salary", "basic pay", "bs"),
    ("employee id", "emp id", "empid", "employee number", "emp no"),
    ("employee code", "emp code", "bio", "bio no", "staff code"),
    ("employee name", "emp name", "empname", "name"),
    (" department ", " dept "),
    ("designation", "desig"),
    ("salary", "pay", "wage"),
    ("conveyance allowance", "conveyance", "conv allowance", "conv allow", "conv allw"),
    ("employer pf", "pf 13%"),
    ("employer esic", "esic 3.25%"),
    ("employee pf", "pf", "pf 12%"),
    ("employee esic", "esic", "esic 0.75%"),
    ("net pay", "net salary"),
    ("bonus", "incentive"),
    ("email", "email address"),
    ("Gender","gen","Gender"),
)

def norm_name(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())

def column_base_name(value):
    return re.sub(
        r"\s+\[duplicate \d+\]$",
        "",
        str(value).strip(),
        flags=re.IGNORECASE,
    )

def column_field_name(value):
    return column_base_name(value).rsplit("|", 1)[-1].strip()

def filter_comparison_columns(df):
    positions = [
        index
        for index, column in enumerate(df.columns)
        if norm_name(column) not in EXCLUDED_COMPARISON_COLUMNS
    ]
    return df.iloc[:, positions].copy()

def column_alias_group(value):
    normalized = norm_name(column_field_name(value))
    for aliases in COLUMN_ALIAS_GROUPS:
        if normalized in {norm_name(alias) for alias in aliases}:
            return norm_name(aliases[0])
    return None

def column_initialism(value):
    tokens = re.findall(
        r"[a-z0-9]+", column_base_name(value).strip().lower()
    )
    if len(tokens) > 1:
        acronym = "".join(token[0] for token in tokens)
        if len(acronym) > 1:
            return acronym
    return None

def columns_match_by_alias_or_initialism(first, second):
    first_group = column_alias_group(first)
    second_group = column_alias_group(second)
    if first_group is not None and first_group == second_group:
        return True

    first_name = norm_name(column_base_name(first))
    second_name = norm_name(column_base_name(second))
    return (
        column_initialism(first) == second_name
        or column_initialism(second) == first_name
    )

def clean_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value

def column_mapping_pairs(mapping):
    """Expand one-to-one and one-to-many column mapping values."""
    for first_column, second_columns in mapping.items():
        if isinstance(second_columns, (list, tuple)):
            for second_column in second_columns:
                yield first_column, second_column
        else:
            yield first_column, second_columns

def values_equal(a, b):
    a, b = clean_value(a), clean_value(b)
    if a is None or b is None:
        return a is b

    def date_like(value):
        if isinstance(value, (pd.Timestamp, datetime.datetime, datetime.date)):
            return pd.Timestamp(value).normalize()
        if isinstance(value, str):
            text = value.strip()
            if re.fullmatch(r"\d{4}-\d{1,2}-\d{1,2}(?:[ T].*)?", text):
                try:
                    return pd.to_datetime(text, errors="raise").normalize()
                except (TypeError, ValueError, OverflowError):
                    pass
        return None

    date_a, date_b = date_like(a), date_like(b)
    if date_a is not None or date_b is not None:
        return date_a is not None and date_b is not None and date_a == date_b

    try:
        return math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-9)
    except (TypeError, ValueError):
        return str(a).strip().casefold() == str(b).strip().casefold()

def build_column_mapping(df1, df2, suggested=None):
    """Map each first-file column to one or more second-file columns or None."""
    cols1, cols2 = list(df1.columns), list(df2.columns)
    mapping = {}
    used = set()

    # Exact names take priority over aliases and model suggestions.
    normalized2 = {}
    for c in cols2:
        normalized2.setdefault(norm_name(column_base_name(c)), []).append(c)
    for c1 in cols1:
        candidates = [
            c for c in normalized2.get(
                norm_name(column_base_name(c1)), []
            )
            if c not in used
        ]
        if candidates:
            mapping[c1] = candidates[0] if len(candidates) == 1 else candidates
            used.update(candidates)

    # Map fields to every equivalent subheader beneath a grouped heading.
    for first_column in cols1:
        first_field = norm_name(column_field_name(first_column))
        first_alias_group = column_alias_group(first_column)
        grouped_candidates = [
            second_column
            for second_column in cols2
            if "|" in str(second_column)
            and (
                norm_name(column_field_name(second_column)) == first_field
                or (
                    first_alias_group is not None
                    and column_alias_group(second_column) == first_alias_group
                )
            )
        ]
        if grouped_candidates:
            existing = mapping.get(first_column)
            mapped_columns = (
                list(existing)
                if isinstance(existing, (list, tuple))
                else [existing] if existing is not None else []
            )
            mapping[first_column] = list(dict.fromkeys(
                mapped_columns + grouped_candidates
            ))
            used.update(grouped_candidates)

    # Match aliases and single-header initialisms only when unambiguous.
    for c1 in cols1:
        if c1 in mapping:
            continue
        candidates = [
            c2 for c2 in cols2
            if c2 not in used and columns_match_by_alias_or_initialism(c1, c2)
        ]
        if len(candidates) == 1:
            mapping[c1] = candidates[0]
            used.add(candidates[0])

    # Apply valid one-to-one LLM suggestions to columns still unmatched.
    for c1, c2 in (suggested or {}).items():
        if c1 in cols1 and c1 not in mapping and c2 in cols2 and c2 not in used:
            mapping[c1] = c2
            used.add(c2)

    for c1 in cols1:
        mapping.setdefault(c1, None)
    return mapping

def find_id_pair(df1, df2, mapping):
    pairs = list(column_mapping_pairs(mapping))
    for c1, c2 in pairs:
        if (
            c2 is not None
            and column_alias_group(c1) == norm_name("employee code")
        ):
            return c1, c2

    id_words = (
        "id", "employeeid", "empid", "staffid", "personid", "recordid",
        "employee number", "employee no",
    )
    for c1, c2 in pairs:
        if c2 is None:
            continue
        n = norm_name(c1)
        if n in {norm_name(x) for x in id_words} or n.endswith("id"):
            return c1, c2
    return None

def find_name_pair(mapping):
    for c1, c2 in column_mapping_pairs(mapping):
        if (
            c2 is not None
            and column_alias_group(c1) == norm_name("employee name")
            and column_alias_group(c2) == norm_name("employee name")
        ):
            return c1, c2
    return None

def unique_text_rows(series):
    rows_by_value = {}
    for index, value in series.items():
        value = clean_value(value)
        if value is not None:
            key = re.sub(r"\s+", " ", str(value).strip()).casefold()
            rows_by_value.setdefault(key, []).append(index)
    return {
        value: indices[0]
        for value, indices in rows_by_value.items()
        if len(indices) == 1
    }

def row_similarity(row1, row2, mapping):
    scores = []
    for c1, c2 in column_mapping_pairs(mapping):
        if c2 is None:
            continue
        a, b = clean_value(row1.get(c1)), clean_value(row2.get(c2))
        if a is None and b is None:
            continue
        if a is None or b is None:
            scores.append(0.0)
        elif values_equal(a, b):
            scores.append(1.0)
        else:
            scores.append(SequenceMatcher(None, str(a).casefold(), str(b).casefold()).ratio())
    return sum(scores) / len(scores) if scores else 0.0

def match_rows(df1, df2, mapping, threshold=0.75):
    """Match by common IDs, unique names, then overall row similarity."""
    pairs, used2 = [], set()
    id_pair = find_id_pair(df1, df2, mapping)
    if id_pair:
        a, b = id_pair
        index2 = {}
        for j, value in df2[b].items():
            v = clean_value(value)
            if v is not None:
                index2.setdefault(str(v).strip().casefold(), []).append(j)
        for i, row in df1.iterrows():
            key = clean_value(row[a])
            candidates = index2.get(str(key).strip().casefold(), []) if key is not None else []
            j = next((x for x in candidates if x not in used2), None)
            if j is not None:
                pairs.append((i, j, 1.0))
                used2.add(j)
            else:
                pairs.append((i, None, 0.0))
        return pairs, set(df1.index) - {x[0] for x in pairs if x[1] is not None}, set(df2.index) - used2, "ID-based"

    # Unique mapped employee names provide a safe key when no shared ID exists.
    name_pair = find_name_pair(mapping)
    used1 = set()
    name_match_count = 0
    if name_pair:
        name1, name2 = name_pair
        unique1 = unique_text_rows(df1[name1])
        unique2 = unique_text_rows(df2[name2])
        for value, i in unique1.items():
            j = unique2.get(value)
            if j is not None:
                pairs.append((i, j, 1.0))
                used1.add(i)
                used2.add(j)
                name_match_count += 1

    # Greedily compare remaining rows using all mapped non-empty fields.
    candidates = []
    for i, r1 in df1.iterrows():
        if i in used1:
            continue
        for j, r2 in df2.iterrows():
            if j in used2:
                continue
            score = row_similarity(r1, r2, mapping)
            if score >= threshold:
                candidates.append((score, i, j))
    candidates.sort(reverse=True)
    for score, i, j in candidates:
        if i not in used1 and j not in used2:
            pairs.append((i, j, score))
            used1.add(i)
            used2.add(j)
    for i in df1.index:
        if i not in used1:
            pairs.append((i, None, 0.0))
    method = (
        "Unique-name + all mapped columns"
        if name_match_count
        else "All mapped columns (no common ID)"
    )
    return pairs, set(df1.index) - used1, set(df2.index) - used2, method


def _build_side_by_side(df1, df2, mapping, row_pairs, unmatched_second):
    headers = []
    fields = []
    first_headers = []
    second_headers = []
    first_indexes = []
    second_indexes = []
    column_pair_indexes = []
    second_column_sources = {}
    for first_column, second_column in column_mapping_pairs(mapping):
        if second_column is not None:
            second_column_sources.setdefault(second_column, []).append(first_column)

    for first_column in df1.columns:
        second_columns = mapping.get(first_column)
        if isinstance(second_columns, (list, tuple)):
            second_columns = list(second_columns)
        elif second_columns is None:
            second_columns = []
        else:
            second_columns = [second_columns]

        if not second_columns:
            continue

        headers.append(f"{first_column}_GH")
        fields.append(("first", first_column))
        first_headers.append(f"{first_column}_GH")
        first_indexes.append(len(headers) - 1)
        first_index = len(headers) - 1

        if second_columns:
            for second_column in second_columns:
                second_header = str(second_column)
                if len(second_column_sources.get(second_column, [])) > 1:
                    second_header = (
                        f"{second_header} ({first_column})"
                    )
                headers.append(f"{second_header}_Client")
                fields.append(("second", second_column))
                second_headers.append(f"{second_header}_Client")
                second_indexes.append(len(headers) - 1)
                column_pair_indexes.append((first_index, len(headers) - 1))
    second_by_first = {i: j for i, j, _ in row_pairs}
    ordered_pairs = [
        (i, second_by_first[i])
        for i in df1.index
    ]
    ordered_pairs.extend(
        (None, j) for j in df2.index if j in unmatched_second
    )

    rows = []
    for first_index, second_index in ordered_pairs:
        values = []
        for source, column in fields:
            if source == "first":
                value = (
                    clean_value(df1.loc[first_index, column])
                    if first_index is not None else None
                )
            else:
                value = (
                    clean_value(df2.loc[second_index, column])
                    if second_index is not None and column is not None else None
                )
            values.append(value)
        rows.append(values)

    side_by_side = pd.DataFrame(rows, columns=headers)
    side_by_side.attrs["first_columns"] = first_headers
    side_by_side.attrs["second_columns"] = second_headers
    side_by_side.attrs["first_indexes"] = first_indexes
    side_by_side.attrs["second_indexes"] = second_indexes
    side_by_side.attrs["column_pair_indexes"] = column_pair_indexes
    return side_by_side


def is_numeric_pair(a, b):
    return (a is not None and b is not None and
            not isinstance(a, (bool,)) and not isinstance(b, (bool,)) and
            pd.api.types.is_number(a) and pd.api.types.is_number(b))

def compare_workbooks(df1, df2, mapping=None, threshold=0.75):
    df1 = filter_comparison_columns(df1)
    df2 = filter_comparison_columns(df2)

    mapping = mapping if mapping is not None else build_column_mapping(df1, df2)
    invalid_first_columns = [column for column in mapping if column not in df1.columns]
    column_pairs = list(column_mapping_pairs(mapping))
    invalid_second_columns = [
        column for _, column in column_pairs
        if column is not None and column not in df2.columns
    ]
    if invalid_first_columns or invalid_second_columns:
        raise ValueError(
            "Column mapping must reference only columns from the uploaded workbooks."
        )
    mapped_second_columns = [
        column for _, column in column_pairs if column is not None
    ]
    # Column-level output: show values from both workbooks for each mapped or missing column.
    col_rows = []
    for c1, c2 in column_pairs:
        col_rows.append({
            "First workbook column": c1,
            "Second workbook column": c2 if c2 else NO_MATCH,
            "Column status": "Matched" if c2 else NO_MATCH,
            "First-file non-empty values": int(df1[c1].notna().sum()),
            "Second-file non-empty values": int(df2[c2].notna().sum()) if c2 else NO_MATCH
        })
    mapped2 = {x for x in mapped_second_columns if x is not None}
    for c2 in df2.columns:
        if c2 not in mapped2:
            col_rows.append({"First workbook column": NO_MATCH, "Second workbook column": c2,
                             "Column status": NO_MATCH,
                             "First-file non-empty values": NO_MATCH,
                             "Second-file non-empty values": int(df2[c2].notna().sum())})

    row_pairs, unmatched1, unmatched2, method = match_rows(
        df1, df2, mapping, threshold
    )
    id_pair = find_id_pair(df1, df2, mapping)
    side_by_side = _build_side_by_side(
        df1, df2, mapping, row_pairs, unmatched2
    )
    side_by_side_column_pairs = [
        (str(c1), str(c2))
        for c1, c2 in column_pairs
        if c2 is not None
    ]
    side_by_side_unmatched_rows = {
        position: "first"
        for position, index in enumerate(df1.index)
        if index in unmatched1
    }
    next_position = len(df1)
    for index in df2.index:
        if index in unmatched2:
            side_by_side_unmatched_rows[next_position] = "second"
            next_position += 1
    output_rows = []
    for i, j, score in row_pairs:
        r1 = df1.loc[i]
        r2 = df2.loc[j] if j is not None else None
        record = {"First-file Excel row": int(i) + 2, "Second-file Excel row": int(j) + 2 if j is not None else NO_MATCH,
                  "Row match method": method, "Row match score": round(score, 3) if j is not None else NO_MATCH}
        for c1, c2 in column_pairs:
            label = f"{c1} | {c2 if c2 else NO_MATCH}"
            v1 = clean_value(r1[c1])
            v2 = clean_value(r2[c2]) if (j is not None and c2 is not None) else None
            record[f"First: {label}"] = v1 if v1 is not None else NO_MATCH
            record[f"Second: {label}"] = v2 if v2 is not None else NO_MATCH
            if j is None or c2 is None:
                record[f"Difference: {label}"] = NO_MATCH
            
            elif is_numeric_pair(v1, v2):
             difference = float(v2) - float(v1)

             if math.isclose(difference, 0, rel_tol=1e-9, abs_tol=1e-9):
              record[f"Difference: {label}"] = "Matched"
             else:
              record[f"Difference: {label}"] = difference

            elif values_equal(v1, v2):
             record[f"Difference: {label}"] = "Matched"

            else:
             record[f"Difference: {label}"] = NO_MATCH

        output_rows.append(record)

    unmatched_rows = []
    for i in sorted(unmatched1):
        unmatched_rows.append({"Source": "First workbook", "Excel row": int(i)+2,
                               **{str(c): clean_value(df1.loc[i, c]) for c in df1.columns}})
    for j in sorted(unmatched2):
        unmatched_rows.append({"Source": "Second workbook", "Excel row": int(j)+2,
                               **{str(c): clean_value(df2.loc[j, c]) for c in df2.columns}})
    matched_count = sum(1 for _, j, _ in row_pairs if j is not None)
    summary = pd.DataFrame([{
        "First workbook rows": len(df1), "Second workbook rows": len(df2),
        "First workbook columns": len(df1.columns),
        "Second workbook columns": len(df2.columns),
        "Matched rows": matched_count, "Unmatched first-file rows": len(unmatched1),
        "Unmatched second-file rows": len(unmatched2), "Row matching method": method,
        "Matched columns": sum(
            1 for _, value in column_pairs if value is not None
        ),
        "Unmatched first-file columns": sum(
            1 for v in mapping.values()
            if v is None or (isinstance(v, (list, tuple)) and not v)
        ),
        "Unmatched second-file columns": len(df2.columns) - len(mapped2)
    }])
    return {"summary": summary, "columns": pd.DataFrame(col_rows),
            "rows": pd.DataFrame(output_rows), "unmatched": pd.DataFrame(unmatched_rows),
            "side_by_side": side_by_side,
            "side_by_side_key_columns": (
                [f"{id_pair[0]}_GH", f"{id_pair[1]}_Client"]
                if id_pair is not None else []
            ),
            "side_by_side_column_pairs": side_by_side_column_pairs,
            "side_by_side_unmatched_rows": side_by_side_unmatched_rows}


def generate_excel_report(result):

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        # Create Excel worksheets
        result["summary"].to_excel(
            writer, sheet_name="Summary", index=False
        )

        result["columns"].to_excel(
            writer, sheet_name="Column Comparison", index=False
        )

        result["rows"].to_excel(
            writer, sheet_name="Row Comparison", index=False
        )

        result["unmatched"].to_excel(
            writer, sheet_name="Unmatched Rows", index=False
        )

        result["side_by_side"].to_excel(
            writer, sheet_name="comparision", index=False
        )

        # Excel colors
        green_fill = PatternFill(
            fill_type="solid", fgColor="C6EFCE"
        )

        yellow_fill = PatternFill(
            fill_type="solid", fgColor="FFEB9C"
        )

        red_fill = PatternFill(
            fill_type="solid", fgColor="FFC7CE"
        )

        green_font = Font(color="006100", bold=True)
        yellow_font = Font(color="9C6500", bold=True)
        red_font = Font(color="9C0006", bold=True)

        # Format Row Comparison worksheet
        ws = writer.book["Row Comparison"]

        for col in range(1, ws.max_column + 1):

            header = ws.cell(row=1, column=col).value

            if header and str(header).startswith("Difference:"):

                for row in range(2, ws.max_row + 1):

                    cell = ws.cell(row=row, column=col)
                    value = cell.value

                    if value == "Matched":
                        cell.fill = green_fill
                        cell.font = green_font

                    elif isinstance(value, (int, float)):

                        if value > 0:
                            cell.fill = yellow_fill
                            cell.font = yellow_font

                        elif value < 0:
                            cell.fill = red_fill
                            cell.font = red_font

                    elif value == "No match found":
                        cell.fill = red_fill
                        cell.font = red_font

        # Highlight unmatched side-by-side rows by their source workbook.
        side_by_side_ws = writer.book["comparision"]
        key_columns = set(result.get("side_by_side_key_columns", []))
        first_fill = PatternFill(fill_type="solid", fgColor="F4CCCC")
        second_fill = PatternFill(fill_type="solid", fgColor="D9D2E9")
        first_numeric_fill = PatternFill(fill_type="solid", fgColor="FCE4D6")
        second_numeric_fill = PatternFill(fill_type="solid", fgColor="DDEBF7")
        key_fill = PatternFill(fill_type="solid", fgColor="C00060")
        key_font = Font(color="FFFFFF", bold=True)
        numeric_columns = {
            index + 1
            for index, dtype in enumerate(result["side_by_side"].dtypes)
            if pd.api.types.is_numeric_dtype(dtype)
        }
        key_column_indexes = {
            cell.column
            for cell in side_by_side_ws[1]
            if cell.value in key_columns
        }
        first_source_indexes = {
            index + 1
            for index in result["side_by_side"].attrs.get("first_indexes", [])
        }
        second_source_indexes = {
            index + 1
            for index in result["side_by_side"].attrs.get("second_indexes", [])
        }

        for data_index, source in result.get(
            "side_by_side_unmatched_rows", {}
        ).items():
            excel_row = data_index + 2
            for column_index in range(1, side_by_side_ws.max_column + 1):
                cell = side_by_side_ws.cell(
                    row=excel_row,
                    column=column_index,
                )
                if cell.value is None:
                    continue
                if column_index in key_column_indexes:
                    cell.fill = key_fill
                    cell.font = key_font
                elif source == "first" and column_index in first_source_indexes:
                    cell.fill = (
                        first_numeric_fill
                        if column_index in numeric_columns else first_fill
                    )
                elif source == "second" and column_index in second_source_indexes:
                    cell.fill = (
                        second_numeric_fill
                        if column_index in numeric_columns else second_fill
                    )

        side_by_side_data = result["side_by_side"]
        unmatched_data_rows = set(
            result.get("side_by_side_unmatched_rows", {})
        )
        yellow_fill = PatternFill(fill_type="solid", fgColor="FFD700")
        for data_index in range(len(side_by_side_data)):
            if data_index in unmatched_data_rows:
                continue
            for first_index, second_index in side_by_side_data.attrs.get(
                "column_pair_indexes", []
            ):
                first_value = side_by_side_data.iloc[data_index, first_index]
                second_value = side_by_side_data.iloc[data_index, second_index]
                if (
                    clean_value(first_value) is not None
                    and clean_value(second_value) is not None
                    and not values_equal(first_value, second_value)
                ):
                    side_by_side_ws.cell(
                        row=data_index + 2,
                        column=first_index + 1,
                    ).fill = yellow_fill
                    side_by_side_ws.cell(
                        row=data_index + 2,
                        column=second_index + 1,
                    ).fill = yellow_fill

        # Format all worksheets
        for sheet in writer.book.worksheets:

            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
            for cell in sheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

            for column_cells in sheet.columns:

                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )

                sheet.column_dimensions[
                    column_cells[0].column_letter
                ].width = min(max_length + 3, 40)

    output.seek(0)
    return output.getvalue()

