#!/usr/bin/env python3
"""diff_core 的基础单元测试：python3 -m pytest test_diff_core.py（或直接运行）。"""

from diff_core import diff_pair, build_diff_workbook


def _ops(segments):
    return [op for _, op in segments]


def test_equal():
    a_seg, b_seg, sim = diff_pair("hello", "hello")
    assert sim == 1.0
    assert set(_ops(a_seg)) == {"equal"}
    assert set(_ops(b_seg)) == {"equal"}


def test_replace_marks_del_and_ins():
    a_seg, b_seg, _ = diff_pair("cat", "cut")
    assert ("a", "del") in a_seg
    assert ("u", "ins") in b_seg


def test_insert_only_on_b():
    a_seg, b_seg, _ = diff_pair("ab", "abc")
    assert set(_ops(a_seg)) == {"equal"}
    assert ("c", "ins") in b_seg


def test_ignore_case():
    _, _, sim = diff_pair("Hello", "hello", ignore_case=True)
    assert sim == 1.0


def test_ignore_whitespace():
    _, _, sim = diff_pair("a b", "ab", ignore_ws=True)
    assert sim == 1.0


def test_word_granularity():
    a_seg, b_seg, _ = diff_pair("the quick fox", "the slow fox", granularity="word")
    assert ("quick", "del") in a_seg
    assert ("slow", "ins") in b_seg


def test_build_workbook_stats():
    pairs = [("a", "a"), ("b", "c"), ("x", "x")]
    wb, result = build_diff_workbook(pairs)
    assert result["total"] == 3
    assert result["diff_count"] == 1
    assert len(result["preview"]) == 3


def test_only_diff_filters_preview():
    pairs = [("a", "a"), ("b", "c")]
    wb, result = build_diff_workbook(pairs, only_diff=True)
    assert len(result["preview"]) == 1
    assert result["preview"][0]["idx"] == 2


if __name__ == "__main__":
    import sys
    fns = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for fn in fns:
        fn()
        print(f"ok  {fn.__name__}")
    print(f"\n全部 {len(fns)} 个测试通过")
    sys.exit(0)
