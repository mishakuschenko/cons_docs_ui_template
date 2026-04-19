from typing import Any

def validate(data: Any) -> float:
    if isinstance(data, (int, float)):
        return float(data)

    if isinstance(data, str):
        cleaned_value = data.replace(" ", "").replace("_", "").replace("-", "")
        try:
            return float(cleaned_value)
        except ValueError:
            raise ValueError()

    try:
        return float(data)
    except (TypeError, ValueError):
        raise ValueError()

def format_input(nums: float) -> str:
    return f"{nums:,.2f}".replace(",", " ")

def unformat_input(input: str) -> float:
    cleaned = input.replace(" ", "")
    cleaned = cleaned.replace(",", ".")
    return float(cleaned)
