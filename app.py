
import os
import streamlit as st
import pandas as pd

from comparator import (
    compare_workbooks,
    generate_excel_report,
    build_column_mapping
)
from bedrock_mapper import suggest_column_mapping


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

/* Entire application background */
.stApp {
    background-color: #eaf2f8;
}

/* Main content background */
[data-testid="stAppViewContainer"] {
    background-color: #eaf2f8;
}

/* Main content spacing */
.block-container {
    padding-top: 2rem;
}

/* Sidebar background */
[data-testid="stSidebar"] {
    background-color: #d6e6f5;
}

/* Header background */
[data-testid="stHeader"] {
    background-color: #eaf2f8;
}

/* Main title banner */
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
    '<div class="hero">'
    '<h1>📊 Excel Smart Comparator</h1>'
    '<p>Compare two workbooks, identify matching columns and rows, '
    'calculate numeric differences and export an audit-ready Excel report.</p>'
    '</div>',
    unsafe_allow_html=True
)


# --------------------------------------------------
# SIDEBAR SETTINGS
# --------------------------------------------------

with st.sidebar:

    st.header("Settings")

    use_bedrock = st.toggle(
        "Use AWS Bedrock for column matching",
        value=True
    )

    st.caption(
        "If disabled or unavailable, normalized column names "
        "are matched automatically."
    )

    model_id = st.text_input(
        "Bedrock model ID",
        value="amazon.nova-micro-v1:0"
    )

    region = st.text_input(
        "AWS region",
        value=os.getenv("AWS_DEFAULT_REGION", "us-east-1")
    )

    threshold = st.slider(
        "Row match threshold",
        0.50,
        1.00,
        0.75,
        0.05,
        help="Used when there is no common ID column. "
             "Matching uses all mapped shared columns."
    )

    st.divider()

    st.caption(
        "Credentials are read securely through the standard "
        "AWS credential chain."
    )


# --------------------------------------------------
# UPLOAD EXCEL FILES
# --------------------------------------------------

c1, c2 = st.columns(2)

with c1:
    old_file = st.file_uploader(
        "Upload first Excel file (baseline)",
        type=["xlsx", "xlsm", "xls"]
    )

with c2:
    new_file = st.file_uploader(
        "Upload second Excel file (new version)",
        type=["xlsx", "xlsm", "xls"]
    )


# --------------------------------------------------
# MAIN APPLICATION
# --------------------------------------------------

if old_file and new_file:

    try:

        # Read Excel workbooks
        xls1 = pd.ExcelFile(old_file)
        xls2 = pd.ExcelFile(new_file)

        s1, s2 = st.columns(2)

        with s1:
            sheet1 = st.selectbox(
                "First workbook sheet",
                xls1.sheet_names
            )

        with s2:
            sheet2 = st.selectbox(
                "Second workbook sheet",
                xls2.sheet_names
            )

        df1 = pd.read_excel(
            old_file,
            sheet_name=sheet1
        )

        df2 = pd.read_excel(
            new_file,
            sheet_name=sheet2
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

            st.dataframe(
                df1.head(8),
                use_container_width=True
            )

        with p2:

            st.metric(
                "Second workbook",
                f"{len(df2):,} rows · {len(df2.columns)} columns"
            )

            st.dataframe(
                df2.head(8),
                use_container_width=True
            )


        # --------------------------------------------------
        # COLUMN MAPPING
        # --------------------------------------------------

        st.subheader("Column Mapping")

        mapping = None

        if use_bedrock:

            if st.button("Suggest column matches with Bedrock"):

                with st.spinner(
                    "Asking Bedrock to identify equivalent columns..."
                ):

                    try:

                        mapping = suggest_column_mapping(
                            list(df1.columns),
                            list(df2.columns),
                            model_id,
                            region
                        )

                        st.session_state["bedrock_mapping"] = mapping

                        st.success(
                            "Bedrock column suggestions received."
                        )

                    except Exception as e:

                        st.warning(
                            f"Bedrock could not be used: {e}. "
                            "Falling back to automatic name matching."
                        )

        if "bedrock_mapping" in st.session_state and use_bedrock:
            mapping = st.session_state["bedrock_mapping"]

        # Automatic column mapping
        auto_mapping = build_column_mapping(
            df1,
            df2,
            mapping
        )

        mapping_rows = [
            {
                "First workbook column": a,
                "Matched second column": (
                    b if b else "No match found"
                )
            }
            for a, b in auto_mapping.items()
        ]

        st.dataframe(
            pd.DataFrame(mapping_rows),
            use_container_width=True,
            hide_index=True
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
                "Comparing columns and rows..."
            ):

                # Compare both Excel files
                result = compare_workbooks(
                    df1,
                    df2,
                    mapping=auto_mapping,
                    threshold=threshold
                )

                # Generate formatted Excel report
                excel_file = generate_excel_report(result)

                # Save results in Streamlit session
                st.session_state["report_bytes"] = excel_file

                st.session_state["report_name"] = (
                    "Excel_Comparison_Report.xlsx"
                )

                st.session_state["result"] = result

                st.success(
                    "Excel comparison completed successfully!"
                )


        # --------------------------------------------------
        # DISPLAY RESULTS
        # --------------------------------------------------

        if "result" in st.session_state:

            result = st.session_state["result"]

            st.success("Comparison completed.")

            m1, m2, m3 = st.columns(3)

            m1.metric(
                "Matched rows",
                int(result["summary"].loc[0, "Matched rows"])
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

            tab1, tab2, tab3 = st.tabs(
                [
                    "Row Comparison",
                    "Column Comparison",
                    "Unmatched Rows"
                ]
            )

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
            # DOWNLOAD EXCEL REPORT
            # --------------------------------------------------

            st.download_button(
                "⬇️ Download Comparison Excel",
                data=st.session_state["report_bytes"],
                file_name=st.session_state["report_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
        "The first file is treated as the baseline and "
        "the second as the updated workbook."
    )
