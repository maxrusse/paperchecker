import pytest

from paperchecker_utils import (
    dedupe_decisions,
    json_pointer_get,
    json_pointer_set,
    normalize_pmid,
)


def test_json_pointer_get_handles_lists_and_invalid_indexes():
    data = {"items": [{"value": 10}, {"value": 20}]}
    assert json_pointer_get(data, "/items/1/value") == 20
    assert json_pointer_get(data, "/items/2/value") is None
    assert json_pointer_get(data, "/items/not-an-index/value") is None


def test_json_pointer_set_rejects_invalid_indexes():
    data = {"items": ["a", "b"]}
    json_pointer_set(data, "/items/1", "c")
    assert data["items"][1] == "c"
    with pytest.raises(ValueError):
        json_pointer_set(data, "/items/3", "d")


def test_normalize_pmid_matches_numeric_variants():
    assert normalize_pmid(123456) == "123456"
    assert normalize_pmid(123456.0) == "123456"
    assert normalize_pmid(" 123456 ") == "123456"
    assert normalize_pmid("123456.0") == "123456"


def test_dedupe_decisions_prefers_latest():
    task_results = [
        {"decisions": [{"path": "/a", "value": 1}]},
        {"decisions": [{"path": "/b", "value": 2}]},
        {"decisions": [{"path": "/a", "value": 3}]},
    ]
    deduped = dedupe_decisions(task_results)
    assert [d["path"] for d in deduped] == ["/b", "/a"]
    assert next(d for d in deduped if d["path"] == "/a")["value"] == 3
