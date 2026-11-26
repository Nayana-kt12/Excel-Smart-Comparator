
import hashlib
import os
import hashlib
from io import BytesIO


import streamlit as st
import pandas as pd

from comparator import (
    compare_workbooks,
    generate_excel_report,
    build_column_mapping,
    column_mapping_pairs,
)

from bedrock_mapper import suggest_column_mapping

from excel_reader import (
    read_sheet_columns,
    read_sheet_with_detected_header,
)


# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------

st.set_page_config(
    page_title="Excel Smart Comparator",
    page_icon="📊",
    layout="wide"
)


# --------------------------------------------------
# CUSTOM UI DESIGN
# --------------------------------------------------

st.markdown("""
<style>

.stApp {
    background-color: #eaf2f8;
}

[data-testid="stAppViewContainer"] {
    background-color: #eaf2f8;
}

.block-container {
    padding-top: 2rem;
}

[data-testid="stSidebar"] {
    background-color: #d6e6f5;
}

[data-testid="stHeader"] {
    background-color: #eaf2f8;
}

.hero {
    background: linear-gradient(120deg, #123b63, #247ba0);
    color: white;
    padding: 22px 26px;
    border-radius: 16px;
    margin-bottom: 20px;
}

.hero h1 {
    color: white;
}

.hero p {
    color: #eaf5ff;
}

</style>
""", unsafe_allow_html=True)


st.markdown(
    """
    <div class="hero">
        <h1>📊 Excel Smart Comparator</h1>
        <p>
        Compare two workbooks, identify matching columns
        and rows, calculate numeric differences and
        export an audit-ready Excel report.
        </p>
    </div>
    """,
    unsafe_allow_html=True
)


# --------------------------------------------------
# SIDEBAR SETTINGS
# --------------------------------------------------

with st.sidebar:

    st.header("Settings")

    use_bedrock = st.toggle(
        "Use Claude Sonnet via AWS Bedrock",
        value=True
    )

    st.caption(
        "Claude AI identifies abbreviations and "
        "semantically equivalent column names."
    )

    model_id = st.text_input(
        "Claude Sonnet Bedrock model ID",
        value="global.anthropic.claude-sonnet-5-5"
    )

    region = st.text_input(
        "AWS region",
        value=os.getenv(
            "AWS_DEFAULT_REGION",
            "ap-south-1"
        )
    )

    threshold = st.slider(
        "Row match threshold",
        min_value=0.50,
        max_value=1.00,
        value=0.75,
        step=0.05,
        help=(
            "Used when there is no common ID column. "
            "Controls row similarity matching."
        )
    )

    st.divider()

    st.caption(
        "AWS credentials are read using the standard "
        "AWS credential chain."
    )


# --------------------------------------------------
# FILE UPLOAD
# --------------------------------------------------

col1, col2 = st.columns(2)

with col1:

    old_file = st.file_uploader(
        "Upload first Excel file (baseline)",
        type=["xlsx", "xlsm", "xls"],
        key="old_excel"
    )

with col2:

    new_file = st.file_uploader(
        "Upload second Excel file (new version)",
        type=["xlsx", "xlsm", "xls"],
        key="new_excel"
    )


# --------------------------------------------------
# MAIN APPLICATION
# --------------------------------------------------

if old_file is not None and new_file is not None:

    try:

        # Read workbooks
        old_bytes = old_file.getvalue()
        new_bytes = new_file.getvalue()

        old_engine_options = (
            {"engine_kwargs": {"read_only": False}}
            if old_file.name.lower().endswith((".xlsx", ".xlsm"))
            else {}
        )
        new_engine_options = (
            {"engine_kwargs": {"read_only": False}}
            if new_file.name.lower().endswith((".xlsx", ".xlsm"))
            else {}
        )
        xls1 = pd.ExcelFile(BytesIO(old_bytes), **old_engine_options)
        xls2 = pd.ExcelFile(BytesIO(new_bytes), **new_engine_options)

        # --------------------------------------------------
        # SHEET SELECTION
        # --------------------------------------------------

        st.subheader("Select Worksheets")

        s1, s2 = st.columns(2)

        with s1:

            sheet1 = st.selectbox(
                "First workbook sheet",
                xls1.sheet_names,
                key="sheet1"
            )

        df1 = read_sheet_with_detected_header(
            xls1,
            sheet1
        )

        sheet_match_scores = {}
        for candidate_sheet in xls2.sheet_names:
            candidate_columns = read_sheet_columns(
                xls2,
                candidate_sheet
            )
            candidate_frame = pd.DataFrame(columns=candidate_columns)
            candidate_mapping = build_column_mapping(
                df1,
                candidate_frame
            )
            sheet_match_scores[candidate_sheet] = sum(
                value is not None
                for value in candidate_mapping.values()
            )

        recommended_sheet = max(
            xls2.sheet_names,
            key=lambda name: sheet_match_scores[name]
        )
        sheet2_key = "sheet2_" + hashlib.sha256(
            old_bytes + b"\0" + new_bytes + sheet1.encode("utf-8")
        ).hexdigest()[:12]

        with s2:

            sheet2 = st.selectbox(
                "Second workbook sheet",
                xls2.sheet_names,
                index=xls2.sheet_names.index(recommended_sheet),
                key=sheet2_key
            )

        st.caption(
            f"Recommended second sheet: {recommended_sheet} "
            f"({sheet_match_scores[recommended_sheet]} matching headers)."
        )

        df2 = read_sheet_with_detected_header(
            xls2,
            sheet2
        )

        # --------------------------------------------------
        # WORKBOOK PREVIEW
        # --------------------------------------------------

        st.subheader("Workbook Preview")

        p1, p2 = st.columns(2)

        with p1:

            st.metric(
                "First workbook",
                f"{len(df1):,} rows · {len(df1.columns)} columns"
            )

            st.caption(
                f"Selected sheet: {sheet1} · Detected header: "
                f"Excel row {df1.attrs.get('detected_header_row', 'not found')}"
            )

            st.dataframe(
                df1.head(8),
                use_container_width=True
            )

        with p2:

            st.metric(
                "Second workbook",
                f"{len(df2):,} rows · {len(df2.columns)} columns"
            )

            st.caption(
                f"Selected sheet: {sheet2} · Detected header: "
                f"Excel row {df2.attrs.get('detected_header_row', 'not found')}"
            )

            st.dataframe(
                df2.head(8),
                use_container_width=True
            )

        # --------------------------------------------------
        # COLUMN MAPPING
        # --------------------------------------------------

        st.subheader("Column Mapping")

        # Unique mapping key for selected workbooks and sheets
        file_hash = hashlib.md5(
            old_bytes + new_bytes
        ).hexdigest()

        mapping_key = (
            f"claude_mapping_{file_hash}_{sheet1}_{sheet2}"
        )
        result_key = f"{mapping_key}_excluding-doj-esic-uan-v2"

        # Clear AI mapping if Bedrock is disabled
        if not use_bedrock:

            st.session_state.pop(
                mapping_key,
                None
            )

        # --------------------------------------------------
        # CLAUDE AI MAPPING
        # --------------------------------------------------

        if use_bedrock:

            st.info(
                "Claude Sonnet AI will analyze column names "
                "and identify abbreviations and equivalent fields."
            )

            if st.button(
                "Suggest column matches with Claude Sonnet",
                type="primary",
                use_container_width=True
            ):

                try:

                    with st.spinner(
                        "Claude Sonnet is analyzing column names..."
                    ):

                        ai_mapping = suggest_column_mapping(
                            old_columns=list(df1.columns),
                            new_columns=list(df2.columns),
                            model_id=model_id.strip(),
                            region=region.strip(),
                            df_old=df1,
                            df_new=df2
                        )

                        st.session_state[mapping_key] = ai_mapping

                    st.success(
                        f"Claude identified {len(ai_mapping)} column matches."
                    )

                except Exception as e:

                    st.session_state.pop(
                        mapping_key,
                        None
                    )

                    st.error(
                        f"Claude Bedrock error: {e}"
                    )

        # --------------------------------------------------
        # AUTOMATIC + AI MAPPING
        # --------------------------------------------------

        ai_mapping = st.session_state.get(
            mapping_key,
            {}
        )

        auto_mapping = build_column_mapping(
            df1,
            df2,
            suggested=ai_mapping
        )

        mapping_rows = []

        for old_column, new_column in column_mapping_pairs(auto_mapping):

            mapping_rows.append({

                "First workbook column": old_column,

                "Second workbook column": (
                    new_column
                    if new_column is not None
                    else "No match found"
                ),

                "Status": (
                    "Matched"
                    if new_column is not None
                    else "No match found"
                ),

                "Mapping source": (
                    "Claude AI"
                    if old_column in ai_mapping
                    and ai_mapping[old_column] == new_column
                    else "Automatic matching"
                )

            })

        # Show columns that exist only in second workbook
        mapped_columns = {
            value
            for _, value in column_mapping_pairs(auto_mapping)
            if value is not None
        }

        for column in df2.columns:

            if column not in mapped_columns:

                mapping_rows.append({

                    "First workbook column": "No match found",

                    "Second workbook column": column,

                    "Status": "Unmatched",

                    "Mapping source": "-"

                })

        st.dataframe(
            pd.DataFrame(mapping_rows),
            use_container_width=True,
            hide_index=True
        )

        st.caption(
            "Review the column mapping before generating "
            "the comparison report."
        )

        # --------------------------------------------------
        # COMPARE WORKBOOKS
        # --------------------------------------------------

        if st.button(
            "Compare workbooks and generate Excel report",
            type="primary",
            use_container_width=True
        ):

            with st.spinner(
                "Comparing Excel columns and rows..."
            ):

                result = compare_workbooks(
                    df1,
                    df2,
                    mapping=auto_mapping,
                    threshold=threshold
                )

                excel_file = generate_excel_report(result)

                st.session_state["report_bytes"] = excel_file

                st.session_state["report_name"] = (
                    "Excel_Comparison_Report.xlsx"
                )

                st.session_state["result"] = result

                st.session_state["result_key"] = result_key

                st.success(
                    "Excel comparison completed successfully!"
                )

        # --------------------------------------------------
        # DISPLAY RESULTS
        # --------------------------------------------------

        if (
            "result" in st.session_state
            and st.session_state.get("result_key") == result_key
        ):

            result = st.session_state["result"]

            st.subheader("Comparison Summary")

            m1, m2, m3 = st.columns(3)

            m1.metric(
                "Matched rows",
                int(
                    result["summary"].loc[
                        0, "Matched rows"
                    ]
                )
            )

            m2.metric(
                "Unmatched first-file rows",
                int(
                    result["summary"].loc[
                        0, "Unmatched first-file rows"
                    ]
                )
            )

            m3.metric(
                "Unmatched second-file rows",
                int(
                    result["summary"].loc[
                        0, "Unmatched second-file rows"
                    ]
                )
            )

            # --------------------------------------------------
            # RESULT TABS
            # --------------------------------------------------

            tab1, tab2, tab3 = st.tabs([
                "Row Comparison",
                "Column Comparison",
                "Unmatched Rows"
            ])

            with tab1:

                st.dataframe(
                    result["rows"],
                    use_container_width=True,
                    height=420
                )

            with tab2:

                st.dataframe(
                    result["columns"],
                    use_container_width=True
                )

            with tab3:

                st.dataframe(
                    result["unmatched"],
                    use_container_width=True,
                    height=350
                )

            # --------------------------------------------------
            # DOWNLOAD REPORT
            # --------------------------------------------------

            st.download_button(
                "⬇️ Download Comparison Excel",
                data=st.session_state["report_bytes"],
                file_name=st.session_state["report_name"],
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
                use_container_width=True
            )

    except Exception as e:

        st.error(
            f"Could not read or compare these workbooks: {e}"
        )

else:

    st.info(
        "Upload two Excel files to begin. "
        "The first file is treated as the baseline "
        "and the second as the updated workbook."
    )

