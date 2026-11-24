
import boto3
import json
import re
import pandas as pd
from botocore.exceptions import ClientError, BotoCoreError


MODEL_ID = "global.anthropic.claude-sonnet-5-5"
AWS_REGION = "ap-south-1"


def normalize_column(name):
    """Remove spaces, punctuation and capitalization differences."""
    return re.sub(
        r"[^a-z0-9]",
        "",
        str(name).casefold()
    )


def get_column_context(df, column):
    """
    Send column name and data type only.
    Do not send actual employee or salary values to the AI.
    """
    if df is None or column not in df.columns:
        return {
            "name": str(column),
            "data_type": "unknown"
        }

    return {
        "name": str(column),
        "data_type": str(df[column].dtype)
    }


def extract_json(text):
    """Extract JSON even if Claude includes code fences."""

    text = text.strip()

    text = re.sub(r"^```(?:json)?\s*", "", text)
    text = re.sub(r"\s*```$", "", text)

    start = text.find("{")

    if start == -1:
        raise ValueError("Claude returned no JSON object.")

    result, _ = json.JSONDecoder().raw_decode(text[start:])

    return result


def suggest_column_mapping(
    old_columns,
    new_columns,
    model_id=MODEL_ID,
    region=AWS_REGION,
    df_old=None,
    df_new=None
):
    """
    Use Anthropic Claude Sonnet through AWS Bedrock
    to identify equivalent Excel column names.

    Returns:
        {
            "Old column": "New column"
        }
    """

    client = boto3.client(
        "bedrock-runtime",
        region_name=region
    )

    old_details = [
        get_column_context(df_old, col)
        for col in old_columns
    ]

    new_details = [
        get_column_context(df_new, col)
        for col in new_columns
    ]

    prompt = f"""Match equivalent Excel column headers. Consider abbreviations, spelling,
case, punctuation, semantic meaning, and the supplied data types. Examples:
Department/DEPT, Basic/BS, Employee Name/Name, Date of Birth/DOB.
Only return confident one-to-one matches. Do not invent headers or include
explanations, confidence scores, or unmatched columns.

First workbook: {json.dumps(old_details, ensure_ascii=False)}
Second workbook: {json.dumps(new_details, ensure_ascii=False)}

Return ONLY a compact JSON object mapping exact first-workbook header strings
to exact second-workbook header strings, e.g. {{"Department":"DEPT"}}.
"""

    try:
        response = client.converse(
            modelId=model_id,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"text": prompt}
                    ]
                }
            ],
            inferenceConfig={
                "maxTokens": 3000
            }
        )

        if response.get("stopReason") == "max_tokens":
            raise ValueError(
                "Claude's mapping response was truncated. Reduce the number "
                "of columns and try again."
            )

        response_text = "".join(
            block["text"]
            for block in response["output"]["message"]["content"]
            if "text" in block
        )

        result = extract_json(response_text)

        if not isinstance(result, dict):
            raise ValueError("Claude's response must be a JSON object.")

        if isinstance(result.get("matches"), list):
            matches = result["matches"]
        else:
            matches = [
                {
                    "old_column": old_name,
                    "new_column": new_name,
                    "confidence": 1.0,
                }
                for old_name, new_name in result.items()
                if isinstance(old_name, str) and isinstance(new_name, str)
            ]

        valid_mapping = {}

        used_old = set()
        used_new = set()

        old_lookup = {
            normalize_column(c): c
            for c in old_columns
        }

        new_lookup = {
            normalize_column(c): c
            for c in new_columns
        }

        for item in matches:

            old_name = item.get("old_column")
            new_name = item.get("new_column")

            confidence = float(
                item.get("confidence", 1.0)
            )

            # Validate returned column names
            old_key = normalize_column(old_name)
            new_key = normalize_column(new_name)

            if old_key not in old_lookup:
                continue

            if new_key not in new_lookup:
                continue

            old_name = old_lookup[old_key]
            new_name = new_lookup[new_key]

            # Reject low-confidence matches
            if confidence < 0.75:
                continue

            # Enforce one-to-one mapping
            if old_name in used_old or new_name in used_new:
                continue

            valid_mapping[old_name] = new_name

            used_old.add(old_name)
            used_new.add(new_name)

        return valid_mapping

    except (ClientError, BotoCoreError):
        raise

    except (ValueError, KeyError, TypeError, json.JSONDecodeError) as e:
        raise ValueError(
            f"Could not process Claude's mapping response: {e}"
        ) from e

