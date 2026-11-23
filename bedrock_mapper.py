import json
import boto3

def suggest_column_mapping(columns1, columns2, model_id, region):
    """Use Bedrock Converse to map semantically equivalent column names.
    Returns a JSON dictionary {first_file_column: second_file_column}.
    """
    client = boto3.client("bedrock-runtime", region_name=region)
    prompt = f"""You are matching spreadsheet column headers.
Return ONLY valid JSON object mapping a header from the first list to its equivalent
header from the second list. Include only confident one-to-one matches. Do not invent
headers. Do not map unrelated fields. Examples: 'Emp Name' -> 'Employee_Name',
'Basic Pay' -> 'Base Salary'. If there is no confident match, omit that key.

First workbook headers: {json.dumps([str(x) for x in columns1], ensure_ascii=False)}
Second workbook headers: {json.dumps([str(x) for x in columns2], ensure_ascii=False)}
"""
    response = client.converse(
        modelId=model_id,
        messages=[{"role": "user", "content": [{"text": prompt}]}],
        inferenceConfig={"temperature": 0, "maxTokens": 1500}
    )
    text = response["output"]["message"]["content"][0]["text"].strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[1].rsplit("```", 1)[0].strip()
    result = json.loads(text)
    if not isinstance(result, dict):
        raise ValueError("Bedrock response was not a JSON mapping.")
    return result
