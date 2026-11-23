# Excel Smart Comparator

A Streamlit application to compare any two Excel workbooks/sheets and export a new Excel report.

## Features
- Upload two Excel files and choose a worksheet from each.
- Map columns using AWS Bedrock semantic matching (optional) with normalized-name matching as a fallback.
- Display columns that do not have a counterpart as **No match found**.
- Match rows automatically:
  - If both sheets have a common ID-like column, use that ID as the authoritative row key.
  - If no common ID exists, score candidate rows using all mapped columns and match one-to-one above the selected threshold.
  - A missing/blank ID or an ID absent from the other workbook is not treated as a match.
- Show values side by side.
- Numeric differences are calculated as **second workbook value − first workbook value**.
- Categorical values show **Match** when equal; otherwise **No match found**.
- Export Summary, Column Comparison, Row Comparison, and Unmatched Rows sheets to a new `.xlsx` file.

## Setup (Windows)

```bash
py -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
streamlit run app.py
```

## AWS Bedrock setup

1. In AWS, enable model access for a model available in your Bedrock region.
2. Configure credentials using one of the standard AWS methods. For example, in PowerShell:

```powershell
$env:AWS_ACCESS_KEY_ID="YOUR_ACCESS_KEY"
$env:AWS_SECRET_ACCESS_KEY="YOUR_SECRET_KEY"
$env:AWS_DEFAULT_REGION="us-east-1"
```

If your organization uses temporary credentials, also set `AWS_SESSION_TOKEN`. Alternatively, configure `aws configure`, use an AWS profile, or run the app on an IAM role with Bedrock permissions.

3. Ensure the identity has permission for `bedrock:InvokeModel` / `bedrock:InvokeModelWithResponseStream` as required by your environment. The app uses the Bedrock Runtime `Converse` API.
4. In the sidebar, enable Bedrock and enter a model ID that is enabled in your account and region. The default shown is an example; replace it if your account uses a different model.

**Security:** Do not paste AWS access keys into the app or commit them to source control. Credentials are read by boto3 from the standard AWS credential chain.

## Important matching behavior

- The application treats a common ID column as the strongest key. It does not silently pair different IDs just because other fields look similar.
- If no common ID column is detected, row similarity is computed over all mapped columns. The threshold is adjustable in the UI. For large sheets, this fallback can be slower because it compares row candidates between both sheets.
- When column names differ semantically (for example, `Emp Name` and `Employee Full Name`), use the Bedrock suggestion button. Review the mapping shown in the UI before generating the report.
- For numerical fields, differences are calculated only when both values are numeric. Text/category mismatches are marked **No match found**.
- A value missing from one side is displayed as **No match found** in the report.

## Project structure

```text
excel_smart_comparator/
├── app.py
├── comparator.py
├── bedrock_mapper.py
├── requirements.txt
└── README.md
```
