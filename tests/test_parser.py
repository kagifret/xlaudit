"""Tests for xlaudit parser — link detection and volatile detection."""

import pytest

from xlaudit.parser import (
    count_cross_sheet_refs,
    count_volatile,
    detect_external_links,
)


# ── External link detection ─────────────────────────────────────────────────

class TestDetectExternalLinks:
    def test_single_external_link(self):
        formulas = ["=[Budget.xlsx]Sheet1!A1"]
        links = detect_external_links(formulas)
        assert links == {"Budget.xlsx"}

    def test_multiple_external_links(self):
        formulas = [
            "=[Budget.xlsx]Sheet1!A1 + [Sales.xlsx]Data!B2",
            "=SUM([Budget.xlsx]Sheet2!C:C)",
        ]
        links = detect_external_links(formulas)
        assert links == {"Budget.xlsx", "Sales.xlsx"}

    def test_external_link_with_path(self):
        formulas = ["='[C:\\Reports\\Q4.xlsx]Summary'!A1"]
        links = detect_external_links(formulas)
        assert links == {"C:\\Reports\\Q4.xlsx"}

    def test_no_external_links(self):
        formulas = ["=SUM(A1:A10)", "=Sheet1!B2"]
        links = detect_external_links(formulas)
        assert links == set()

    def test_empty_formulas(self):
        assert detect_external_links([]) == set()


# ── Volatile function detection ──────────────────────────────────────────────

class TestCountVolatile:
    def test_now_function(self):
        assert count_volatile(["=NOW()"]) == 1

    def test_today_function(self):
        assert count_volatile(["=TODAY()"]) == 1

    def test_rand_function(self):
        assert count_volatile(["=RAND()"]) == 1

    def test_randbetween_function(self):
        assert count_volatile(["=RANDBETWEEN(1,100)"]) == 1

    def test_indirect_function(self):
        assert count_volatile(["=INDIRECT(A1)"]) == 1

    def test_offset_function(self):
        assert count_volatile(["=OFFSET(A1,1,1)"]) == 1

    def test_multiple_volatile_in_one_formula(self):
        assert count_volatile(["=NOW() + TODAY()"]) == 2

    def test_no_volatile(self):
        assert count_volatile(["=SUM(A1:A10)", "=IF(B1>0,1,0)"]) == 0

    def test_case_insensitive(self):
        assert count_volatile(["=now()", "=Today()", "=RAND()"]) == 3

    def test_not_a_function_call(self):
        # "NOW" without parentheses should not count
        assert count_volatile(["=A1+NOW_COLUMN"]) == 0


# ── Cross-sheet reference detection ─────────────────────────────────────────

class TestCountCrossSheetRefs:
    def test_simple_cross_sheet(self):
        assert count_cross_sheet_refs(["=Sheet1!A1"]) == 1

    def test_quoted_sheet_name(self):
        assert count_cross_sheet_refs(["='My Sheet'!B2"]) == 1

    def test_no_cross_sheet(self):
        assert count_cross_sheet_refs(["=SUM(A1:A10)"]) == 0

    def test_excludes_external_links(self):
        # External links contain [...] so should NOT count as cross-sheet
        formulas = ["=[Budget.xlsx]Sheet1!A1"]
        assert count_cross_sheet_refs(formulas) == 0

    def test_multiple_refs(self):
        formulas = ["=Sheet1!A1 + Sheet2!B2"]
        assert count_cross_sheet_refs(formulas) == 2
