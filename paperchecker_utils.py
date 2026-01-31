import re
from typing import Any, Dict, Iterable, List, Optional

NUMERIC_TOL_ABS = 0.01
NUMERIC_TOL_REL = 0.01


def normalize_string(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    stripped = value.strip()
    return stripped if stripped != "" else None


def normalize_excel_value(value: Any) -> Any:
    if isinstance(value, bool):
        return 1 if value else 0
    if isinstance(value, str):
        return normalize_string(value)
    return value


def coerce_float(value: Any) -> Optional[float]:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        try:
            return float(value.strip())
        except ValueError:
            return None
    return None


def values_match(a: Any, b: Any, abs_tol: float = NUMERIC_TOL_ABS, rel_tol: float = NUMERIC_TOL_REL) -> bool:
    if a is None and b is None:
        return True
    if isinstance(a, str) and isinstance(b, str):
        return a.strip() == b.strip()
    fa = coerce_float(a)
    fb = coerce_float(b)
    if fa is not None and fb is not None:
        if abs(fa - fb) <= abs_tol:
            return True
        if abs(fb) > 0 and abs(fa - fb) / abs(fb) <= rel_tol:
            return True
    return a == b


def normalize_pmid(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()
    if isinstance(value, str):
        stripped = value.strip()
        if stripped == "":
            return None
        if stripped.isdigit():
            return str(int(stripped))
        if stripped.endswith(".0") and stripped[:-2].isdigit():
            return str(int(stripped[:-2]))
        return stripped
    return str(value)


def json_pointer_get(obj: Any, pointer: str) -> Any:
    if pointer in ("", "/"):
        return obj
    parts = pointer.lstrip("/").split("/")
    cur = obj
    for part in parts:
        token = part.replace("~1", "/").replace("~0", "~")
        if isinstance(cur, list):
            if not token.isdigit():
                return None
            idx = int(token)
            if idx < 0 or idx >= len(cur):
                return None
            cur = cur[idx]
        elif isinstance(cur, dict):
            if token not in cur:
                return None
            cur = cur.get(token)
        else:
            return None
    return cur


def json_pointer_set(obj: Any, pointer: str, value: Any) -> None:
    if pointer in ("", "/"):
        raise ValueError("json pointer cannot be empty when setting")
    parts = pointer.lstrip("/").split("/")
    cur = obj
    for idx, part in enumerate(parts):
        token = part.replace("~1", "/").replace("~0", "~")
        last = idx == len(parts) - 1
        if isinstance(cur, list):
            if not token.isdigit():
                raise ValueError(f"json pointer list index is not an integer: {token}")
            list_idx = int(token)
            if list_idx < 0 or list_idx >= len(cur):
                raise ValueError(f"json pointer list index out of range: {list_idx}")
            if last:
                cur[list_idx] = value
            else:
                cur = cur[list_idx]
        elif isinstance(cur, dict):
            if last:
                cur[token] = value
            else:
                if token not in cur or not isinstance(cur[token], (dict, list)):
                    cur[token] = {}
                cur = cur[token]
        else:
            raise ValueError("json pointer target is not a container")


def extract_page_from_evidence(evidence_text: str) -> Optional[int]:
    match = re.search(r"\bPAGE\s+(\d+)\b", evidence_text or "")
    if match:
        try:
            return int(match.group(1))
        except ValueError:
            return None
    return None


def dedupe_decisions(task_results: Iterable[Dict[str, Any]]) -> List[Dict[str, Any]]:
    ordered_paths: List[str] = []
    decisions_by_path: Dict[str, Dict[str, Any]] = {}
    for result in task_results or []:
        for decision in result.get("decisions") or []:
            path = decision.get("path")
            if not path:
                continue
            if path in decisions_by_path:
                ordered_paths = [p for p in ordered_paths if p != path]
            decisions_by_path[path] = decision
            ordered_paths.append(path)
    return [decisions_by_path[p] for p in ordered_paths]
