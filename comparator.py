import re
import math
import pandas as pd
from difflib import SequenceMatcher
from io import BytesIO
from openpyxl.styles import PatternFill, Font, Alignment

NO_MATCH = "No match found"

def norm_name(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())

def clean_value(value):
    if pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value

def values_equal(a, b):
    a, b = clean_value(a), clean_value(b)
    if a is None or b is None:
        return a is b
    try:
        return math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-9)
    except (TypeError, ValueError):
        return str(a).strip().casefold() == str(b).strip().casefold()

def build_column_mapping(df1, df2, suggested=None):
    """Map each first-file column to one second-file column or None."""
    cols1, cols2 = list(df1.columns), list(df2.columns)
    mapping = {}
    used = set()
    suggested = suggested or {}
    # accept only valid, one-to-one suggestions
    for c1, c2 in suggested.items():
        if c1 in cols1 and c2 in cols2 and c2 not in used:
            mapping[c1] = c2
            used.add(c2)
    normalized2 = {}
    for c in cols2:
        normalized2.setdefault(norm_name(c), []).append(c)
    for c1 in cols1:
        if c1 in mapping:
            continue
        candidates = normalized2.get(norm_name(c1), [])
        candidate = next((c for c in candidates if c not in used), None)
        if candidate is not None:
            mapping[c1] = candidate
            used.add(candidate)
        else:
            mapping[c1] = None
    return mapping

def find_id_pair(df1, df2, mapping):
    id_words = ("id", "employeeid", "empid", "staffid", "personid", "recordid", "employee number", "employee no")
    for c1, c2 in mapping.items():
        if c2 is None:
            continue
        n = norm_name(c1)
        if n in {norm_name(x) for x in id_words} or n.endswith("id"):
            return c1, c2
    return None

def row_similarity(row1, row2, mapping):
    scores = []
    for c1, c2 in mapping.items():
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
    """One-to-one row matching. A common ID column is authoritative when present."""
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

    # No shared ID: compare every mapped, non-empty field. Greedy one-to-one best match.
    candidates = []
    for i, r1 in df1.iterrows():
        for j, r2 in df2.iterrows():
            score = row_similarity(r1, r2, mapping)
            if score >= threshold:
                candidates.append((score, i, j))
    candidates.sort(reverse=True)
    used1 = set()
    for score, i, j in candidates:
        if i not in used1 and j not in used2:
            pairs.append((i, j, score))
            used1.add(i)
            used2.add(j)
    for i in df1.index:
        if i not in used1:
            pairs.append((i, None, 0.0))
    return pairs, set(df1.index) - used1, set(df2.index) - used2, "All mapped columns (no common ID)"

def is_numeric_pair(a, b):
    return (a is not None and b is not None and
            not isinstance(a, (bool,)) and not isinstance(b, (bool,)) and
            pd.api.types.is_number(a) and pd.api.types.is_number(b))

def compare_workbooks(df1, df2, mapping=None, threshold=0.75):
    mapping = mapping or build_column_mapping(df1, df2)
    # Column-level output: show values from both workbooks for each mapped or missing column.
    col_rows = []
    for c1, c2 in mapping.items():
        col_rows.append({
            "First workbook column": c1,
            "Second workbook column": c2 if c2 else NO_MATCH,
            "Column status": "Matched" if c2 else NO_MATCH,
            "First-file non-empty values": int(df1[c1].notna().sum()),
            "Second-file non-empty values": int(df2[c2].notna().sum()) if c2 else NO_MATCH
        })
    mapped2 = {x for x in mapping.values() if x is not None}
    for c2 in df2.columns:
        if c2 not in mapped2:
            col_rows.append({"First workbook column": NO_MATCH, "Second workbook column": c2,
                             "Column status": NO_MATCH,
                             "First-file non-empty values": NO_MATCH,
                             "Second-file non-empty values": int(df2[c2].notna().sum())})

    pairs, unmatched1, unmatched2, method = match_rows(df1, df2, mapping, threshold)
    output_rows = []
    for i, j, score in pairs:
        r1 = df1.loc[i]
        r2 = df2.loc[j] if j is not None else None
        record = {"First-file Excel row": int(i) + 2, "Second-file Excel row": int(j) + 2 if j is not None else NO_MATCH,
                  "Row match method": method, "Row match score": round(score, 3) if j is not None else NO_MATCH}
        for c1, c2 in mapping.items():
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
    matched_count = sum(1 for _, j, _ in pairs if j is not None)
    summary = pd.DataFrame([{
        "First workbook rows": len(df1), "Second workbook rows": len(df2),
        "Matched rows": matched_count, "Unmatched first-file rows": len(unmatched1),
        "Unmatched second-file rows": len(unmatched2), "Row matching method": method,
        "Matched columns": sum(1 for v in mapping.values() if v is not None),
        "Unmatched first-file columns": sum(1 for v in mapping.values() if v is None),
        "Unmatched second-file columns": len(df2.columns) - len(mapped2)
    }])
    return {"summary": summary, "columns": pd.DataFrame(col_rows),
            "rows": pd.DataFrame(output_rows), "unmatched": pd.DataFrame(unmatched_rows)}


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

        # Format all worksheets
        for sheet in writer.book.worksheets:

            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions

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

