from io import BytesIO
from pathlib import Path
from zipfile import ZIP_STORED, ZipFile

from deal_expenses.lob_workbooks import LOB_NAMES, _contiguous_ranges, build_output_archive, refresh_workbook_pivots


def test_build_output_archive_places_all_workbooks_in_master_folder(tmp_path: Path):
    output_paths = [tmp_path / "Expense Report Analysis.xlsx"]
    output_paths.extend(tmp_path / f"Expense Report Analysis - {lob_name}.xlsx" for lob_name in LOB_NAMES)
    for index, output_path in enumerate(output_paths):
        output_path.write_bytes(f"workbook-{index}".encode())

    archive = build_output_archive(output_paths, "Expense Report Analysis")

    with ZipFile(BytesIO(archive)) as zip_file:
        assert zip_file.namelist() == [f"Expense Report Analysis/{path.name}" for path in output_paths]
        assert zip_file.read("Expense Report Analysis/Expense Report Analysis.xlsx") == b"workbook-0"
        assert all(entry.compress_type == ZIP_STORED for entry in zip_file.infolist())


def test_contiguous_ranges_groups_adjacent_rows_for_batched_deletion():
    assert _contiguous_ranges([2, 3, 4, 7, 9, 10]) == [(2, 4), (7, 7), (9, 10)]


class StubPivotCache:
    def __init__(self) -> None:
        self.refresh_count = 0

    def Refresh(self) -> None:
        self.refresh_count += 1


class StubPivotTable:
    def __init__(self, cache_index: int, cache: StubPivotCache) -> None:
        self.CacheIndex = cache_index
        self._cache = cache
        self.refresh_count = 0

    def PivotCache(self) -> StubPivotCache:
        return self._cache

    def RefreshTable(self) -> None:
        self.refresh_count += 1


class StubPivotTables:
    def __init__(self, items: list[StubPivotTable]) -> None:
        self._items = items
        self.Count = len(items)

    def Item(self, index: int) -> StubPivotTable:
        return self._items[index - 1]


class StubWorksheet:
    def __init__(self, pivot_tables: StubPivotTables) -> None:
        self._pivot_tables = pivot_tables

    def PivotTables(self) -> StubPivotTables:
        return self._pivot_tables


class StubWorksheets:
    def __init__(self, items: list[StubWorksheet]) -> None:
        self._items = items
        self.Count = len(items)

    def __call__(self, index: int) -> StubWorksheet:
        return self._items[index - 1]


def test_refresh_workbook_pivots_refreshes_shared_caches_once():
    shared_cache = StubPivotCache()
    other_cache = StubPivotCache()
    first_pivot = StubPivotTable(1, shared_cache)
    second_pivot = StubPivotTable(1, shared_cache)
    third_pivot = StubPivotTable(2, other_cache)
    workbook = type("StubWorkbook", (), {"Worksheets": StubWorksheets([StubWorksheet(StubPivotTables([first_pivot, second_pivot, third_pivot]))])})()

    refresh_workbook_pivots(workbook)

    assert shared_cache.refresh_count == 1
    assert other_cache.refresh_count == 1
    assert [pivot.refresh_count for pivot in (first_pivot, second_pivot, third_pivot)] == [1, 1, 1]