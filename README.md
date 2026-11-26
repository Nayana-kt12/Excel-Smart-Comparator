# Excel Smart Comparator

A Streamlit application to compare any two Excel workbooks/sheets and export a new Excel report.

## Features
- Upload two Excel files and choose a worksheet from each. The app detects header rows in the first 20 rows, including sheets with title rows above their headers, and recommends the second workbook sheet with the most matching column headers.
- Map columns using Anthropic Claude Sonnet through AWS Bedrock (optional), with normalized-name, alias, and initialism matching as a fallback.
- Display columns that do not have a counterpart as **No match found**.
- Match rows automatically:
  - If both sheets have a common ID-like column, use that ID as the authoritative row key.
  - If no common ID exists, score candidate rows using all mapped columns and match one-to-one above the selected threshold.
  - A missing/blank ID or an ID absent from the other workbook is not treated as a match.
- Show values side by side.
- Paired columns in the exported side-by-side sheet are labeled with `_GH` and `_Client` suffixes; repeated or equivalent headers use their grouped heading when available, such as `Master Wage | BASIC`. A first-workbook field maps to every matching grouped subheader in the second workbook.
- Show employee-code columns from the first and second workbooks separately.
- The side-by-side sheet includes mapped columns only and highlights differing values, including text and numbers, in yellow. Numeric differences are calculated as **second workbook value − first workbook value**.
- Report worksheet column headers are bold and centered.
- Categorical values show **Match** when equal; otherwise **No match found**.
- Match rows by employee code/ID when a common code or ID column is available.
- Export the side-by-side comparison, comparison results, and column mapping details to a new `.xlsx` file. Original worksheets from the uploaded workbooks are not copied into the report.
- Highlight unmatched side-by-side rows by source, with employee-code cells in dark pink and numeric cells in distinct source colors.

## Setup (Windows)

```bash
py -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
streamlit run app.py
```

## Claude Sonnet via AWS Bedrock

1. In AWS, enable access to Claude Sonnet 5.5 and its global inference profile (`global.anthropic.claude-sonnet-5-5`) for your account and region.
2. Configure credentials using one of the standard AWS methods. For example, in PowerShell:

```powershell
$env:AWS_ACCESS_KEY_ID="YOUR_ACCESS_KEY"
$env:AWS_SECRET_ACCESS_KEY="YOUR_SECRET_KEY"
$env:AWS_DEFAULT_REGION="us-east-1"
```

If your organization uses temporary credentials, also set `AWS_SESSION_TOKEN`. Alternatively, configure `aws configure`, use an AWS profile, or run the app on an IAM role with Bedrock permissions.

3. Ensure the identity has permission for `bedrock:InvokeModel` / `bedrock:InvokeModelWithResponseStream` as required by your environment. The app uses the Bedrock Runtime `Converse` API.
4. In the sidebar, enable Claude Sonnet and enter a model ID or inference profile enabled in your account and region. The app uses the Bedrock Runtime `Converse` API; it does not call the Anthropic API directly.

**Security:** Do not paste AWS access keys into the app or commit them to source control. Credentials are read by boto3 from the standard AWS credential chain.

## Important matching behavior

- The application treats a common ID column as the strongest key. It does not silently pair different IDs just because other fields look similar.
- Common employee code headers such as `Employee Code`, `Emp Code`, and `BIO NO.` are treated as row keys.
- If no common ID column is detected, unique matches on a mapped employee-name column are paired first; duplicate names are not used as keys. Remaining rows are matched by similarity across mapped columns. The threshold is adjustable in the UI.
- Exact normalized names are matched first. Common aliases such as `Department`/`DEPT`, `Basic`/`BS`, `Employee Name`/`Name`, and initialisms of multiword headers are matched automatically when unambiguous.
- When column names differ semantically beyond those aliases, use the Claude Sonnet suggestion button. Review the mapping shown in the UI before generating the report.
- For numerical fields, differences are calculated only when both values are numeric. Text/category mismatches are marked **No match found**.
- A value missing from one side is displayed as **No match found** in the report.

## Project structure

```text
excel_smart_comparator/
├── app.py
├── comparator.py
├── bedrock_mapper.py
├── excel_reader.py
├── requirements.txt
└── README.md
```
