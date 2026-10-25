"""Small helpers for working with Purview response objects."""
import re


def item_name(item):
    """Best available name: terms/domains may use different fields."""
    return str(item.get("name") or item.get("displayName") or item.get("nickName") or item.get("CDE ColumnName") or "-")


def remove_html_tags(text):
    """Descriptions come as '<div>text</div>'; keep only the text."""
    return re.sub(r"<[^>]+>", "", text or "").strip()


def flatten_fields(value, path=""):
    """Yield (field_path, value) for every field, however deeply nested."""
    if isinstance(value, dict) and value:
        for key, item in value.items():
            yield from flatten_fields(item, f"{path}.{key}" if path else key)
    elif isinstance(value, list) and value:
        for index, item in enumerate(value):
            yield from flatten_fields(item, f"{path}[{index}]")
    else:
        yield path, value
