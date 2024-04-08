import re


def flatten_dict(d: dict, parent_key='', sep='_') -> dict:
    """展平字典"""
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)


def match_clean(string: str) -> str:
    """去除字符串中的所有空白/特殊字符"""
    return re.sub(r"\s+", "", string.strip())
