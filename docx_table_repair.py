"""Transactional repair of Word table topology after PDF conversion.

The repair is intentionally conservative. It only groups consecutive Word
cells when geometry, text, and cell styling prove that they belong to one
source PDF cell. The original DOCX is not replaced until the repaired OOXML
has been written and reopened successfully.
"""

from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
import os
from pathlib import Path
import unicodedata
from zipfile import BadZipFile, ZipFile
from uuid import uuid4

from lxml import etree as ElementTree


_WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_MATH_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_MARKUP_COMPATIBILITY_NAMESPACE = (
    "http://schemas.openxmlformats.org/markup-compatibility/2006"
)
_MARKUP_PREFIX_ATTRIBUTES = (
    "Ignorable",
    "MustUnderstand",
    "PreserveAttributes",
    "PreserveElements",
    "ProcessContent",
)


def _qname(name: str) -> str:
    return f"{{{_WORD_NAMESPACE}}}{name}"


_FORBIDDEN_CONTENT_TAGS = frozenset(
    {
        _qname("altChunk"),
        _qname("bookmarkEnd"),
        _qname("bookmarkStart"),
        _qname("commentRangeEnd"),
        _qname("commentRangeStart"),
        _qname("del"),
        _qname("drawing"),
        _qname("fldChar"),
        _qname("fldSimple"),
        _qname("ins"),
        _qname("instrText"),
        _qname("moveFrom"),
        _qname("moveTo"),
        _qname("object"),
        _qname("pict"),
        _qname("sdt"),
        _qname("tbl"),
        f"{{{_MATH_NAMESPACE}}}oMath",
        f"{{{_MATH_NAMESPACE}}}oMathPara",
    }
)


class DocxTableRepairError(RuntimeError):
    """Raised when a table cannot be repaired without guessing."""


@dataclass(frozen=True, slots=True)
class _SourceSegment:
    row: int
    start: int
    end: int
    owner_row: int
    owner_column: int
    row_span: int

    @property
    def continuation(self) -> bool:
        return self.row != self.owner_row


@dataclass(frozen=True, slots=True)
class _OutputCell:
    element: ElementTree.Element
    start: int
    end: int
    desired_width: float
    text: str


@dataclass(frozen=True, slots=True)
class _SourceCellBorderEdges:
    top: bool
    left: bool
    bottom: bool
    right: bool


_CELL_BORDER_EDGE_NAMES = ("top", "left", "bottom", "right")
_CELL_BORDER_EDGE_ALIASES = {
    "top": ("top",),
    "left": ("left", "start"),
    "bottom": ("bottom",),
    "right": ("right", "end"),
}
_COMPLEX_CELL_BORDER_NAMES = frozenset(
    {"insideH", "insideV", "tl2br", "tr2bl"}
)


def _integer_attribute(element, name: str, default=None):
    if element is None:
        return default
    value = element.get(_qname(name))
    if value is None:
        return default
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _normalize_cell_text(value) -> str:
    characters = []
    for character in unicodedata.normalize("NFC", str(value or "")):
        if unicodedata.category(character).startswith("P"):
            compatible = unicodedata.normalize("NFKC", character)
            if (
                len(compatible) == 1
                and unicodedata.category(compatible).startswith("P")
            ):
                character = compatible
        characters.append(character)

    normalized = []
    for index, character in enumerate(characters):
        if character in "\r\n\v\f":
            continue
        if not character.isspace():
            normalized.append(character)
            continue
        previous = next(
            (item for item in reversed(normalized) if not item.isspace()),
            "",
        )
        following = next(
            (
                item
                for item in characters[index + 1 :]
                if not item.isspace()
            ),
            "",
        )
        if (
            previous
            and following
            and previous.isascii()
            and following.isascii()
            and previous.isalnum()
            and following.isalnum()
            and (not normalized or normalized[-1] != " ")
        ):
            normalized.append(" ")
    return "".join(normalized).strip()


def _cell_text(cell) -> str:
    return "".join(
        node.text or "" for node in cell.findall(f".//{_qname('t')}")
    )


def _cell_properties(cell, *, create=False):
    properties = cell.find(f"./{_qname('tcPr')}")
    if properties is None and create:
        properties = ElementTree.Element(_qname("tcPr"))
        cell.insert(0, properties)
    return properties


def _grid_span(cell) -> int:
    properties = _cell_properties(cell)
    span = (
        properties.find(f"./{_qname('gridSpan')}")
        if properties is not None
        else None
    )
    value = _integer_attribute(span, "val", 1)
    return max(1, int(value or 1))


def _vertical_merge_value(cell) -> str:
    properties = _cell_properties(cell)
    merge = (
        properties.find(f"./{_qname('vMerge')}")
        if properties is not None
        else None
    )
    if merge is None:
        return ""
    return str(merge.get(_qname("val"), "continue")).strip().lower()


def _table_grid_widths(table) -> tuple[int, ...]:
    grid = table.find(f"./{_qname('tblGrid')}")
    if grid is None:
        raise DocxTableRepairError("Word table has no tblGrid definition")
    widths = tuple(
        _integer_attribute(column, "w", 0)
        for column in grid.findall(f"./{_qname('gridCol')}")
    )
    if not widths or any(width is None or width <= 0 for width in widths):
        raise DocxTableRepairError("Word table grid widths are incomplete")
    return tuple(int(width) for width in widths)


def _cell_desired_width(cell, grid_widths, start, end) -> float:
    properties = _cell_properties(cell)
    width = (
        properties.find(f"./{_qname('tcW')}")
        if properties is not None
        else None
    )
    width_type = (
        str(width.get(_qname("type"), "dxa")).strip().lower()
        if width is not None
        else ""
    )
    value = _integer_attribute(width, "w", 0)
    if width_type in {"", "dxa"} and value and value > 0:
        return float(value)
    return float(sum(grid_widths[start:end]))


def _output_rows(table, grid_widths) -> tuple[tuple[_OutputCell, ...], ...]:
    column_count = len(grid_widths)
    result = []
    for row in table.findall(f"./{_qname('tr')}"):
        row_properties = row.find(f"./{_qname('trPr')}")
        before = (
            row_properties.find(f"./{_qname('gridBefore')}")
            if row_properties is not None
            else None
        )
        after = (
            row_properties.find(f"./{_qname('gridAfter')}")
            if row_properties is not None
            else None
        )
        start = max(0, int(_integer_attribute(before, "val", 0) or 0))
        cells = []
        for cell in row.findall(f"./{_qname('tc')}"):
            span = _grid_span(cell)
            end = start + span
            if end > column_count:
                raise DocxTableRepairError("Word row extends beyond tblGrid")
            cells.append(
                _OutputCell(
                    element=cell,
                    start=start,
                    end=end,
                    desired_width=_cell_desired_width(
                        cell,
                        grid_widths,
                        start,
                        end,
                    ),
                    text=_cell_text(cell),
                )
            )
            start = end
        start += max(0, int(_integer_attribute(after, "val", 0) or 0))
        if start != column_count or not cells:
            raise DocxTableRepairError("Word row does not cover tblGrid exactly")
        result.append(tuple(cells))
    if not result:
        raise DocxTableRepairError("Word table has no rows")
    return tuple(result)


def _existing_signature(table, grid_widths):
    output_rows = _output_rows(table, grid_widths)
    column_count = len(grid_widths)
    matrices = []
    spans: list[list[int]] = []
    active_merges: dict[tuple[int, int], int] = {}
    for row_index, row in enumerate(output_rows):
        values = [None] * column_count
        current_merges = {}
        for cell in row:
            merge_value = _vertical_merge_value(cell.element)
            key = (cell.start, cell.end - cell.start)
            if merge_value and merge_value != "restart":
                span_index = active_merges.get(key)
                if span_index is None or _normalize_cell_text(cell.text):
                    raise DocxTableRepairError("Invalid vertical merge chain")
                spans[span_index][2] += 1
                current_merges[key] = span_index
                value = None
            else:
                span_index = len(spans)
                spans.append(
                    [row_index, cell.start, 1, cell.end - cell.start]
                )
                value = cell.text
                if merge_value == "restart":
                    current_merges[key] = span_index
            values[cell.start] = value
        matrices.append(tuple(values))
        active_merges = current_merges
    return (
        (len(output_rows), column_count),
        tuple(matrices),
        tuple(tuple(item) for item in sorted(spans)),
    )


def _coerce_source_table(shape, matrix, spans, widths):
    try:
        row_count, column_count = int(shape[0]), int(shape[1])
        matrix = tuple(tuple(row) for row in matrix)
        spans = tuple(tuple(int(value) for value in span) for span in spans)
        widths = tuple(float(width) for width in widths)
    except (TypeError, ValueError, IndexError) as exc:
        raise DocxTableRepairError("Source table topology is malformed") from exc
    if (
        row_count < 1
        or column_count < 1
        or len(matrix) != row_count
        or any(len(row) != column_count for row in matrix)
        or len(widths) != column_count
        or any(width <= 0 for width in widths)
    ):
        raise DocxTableRepairError("Source table dimensions are incomplete")

    coverage = [[None] * column_count for _row in range(row_count)]
    owners = {}
    for span in spans:
        if len(span) != 4:
            raise DocxTableRepairError("Source table span is malformed")
        row, column, row_span, column_span = span
        if (
            row < 0
            or column < 0
            or row_span < 1
            or column_span < 1
            or row + row_span > row_count
            or column + column_span > column_count
        ):
            raise DocxTableRepairError("Source table span is out of range")
        owner = (row, column, row_span, column_span)
        owners[(row, column)] = owner
        for covered_row in range(row, row + row_span):
            for covered_column in range(column, column + column_span):
                if coverage[covered_row][covered_column] is not None:
                    raise DocxTableRepairError("Source table spans overlap")
                coverage[covered_row][covered_column] = owner

    if any(value is None for row in coverage for value in row):
        raise DocxTableRepairError("Source table spans leave a grid gap")
    for row_index, row in enumerate(matrix):
        for column_index, value in enumerate(row):
            is_owner = (row_index, column_index) in owners
            if is_owner == (value is None):
                raise DocxTableRepairError(
                    "Source table matrix does not match its cell spans"
                )

    segments = []
    for row_index, row_coverage in enumerate(coverage):
        row_segments = []
        column = 0
        while column < column_count:
            owner = row_coverage[column]
            end = column + 1
            while end < column_count and row_coverage[end] == owner:
                end += 1
            owner_row, owner_column, row_span, _column_span = owner
            row_segments.append(
                _SourceSegment(
                    row=row_index,
                    start=column,
                    end=end,
                    owner_row=owner_row,
                    owner_column=owner_column,
                    row_span=row_span,
                )
            )
            column = end
        segments.append(tuple(row_segments))
    return (
        (row_count, column_count),
        matrix,
        tuple(sorted(spans)),
        widths,
        tuple(segments),
    )


def _coerce_source_border_edges(
    shape,
    span_order,
    normalized_spans,
    border_edges,
):
    try:
        input_spans = tuple(
            tuple(int(value) for value in span) for span in span_order
        )
        input_edges = tuple(tuple(edge) for edge in border_edges)
    except (TypeError, ValueError) as exc:
        raise DocxTableRepairError(
            "Source table border evidence is malformed"
        ) from exc
    if len(input_edges) != len(input_spans):
        raise DocxTableRepairError(
            "Source table border evidence does not match its cell spans"
        )

    edges_by_span = {}
    for span, values in zip(input_spans, input_edges):
        if len(values) != 4 or any(type(value) is not bool for value in values):
            raise DocxTableRepairError(
                "Source cell borders must be four strict booleans"
            )
        if span in edges_by_span:
            raise DocxTableRepairError(
                "Source table border evidence contains a duplicate cell span"
            )
        edges_by_span[span] = _SourceCellBorderEdges(*values)
    if set(edges_by_span) != set(normalized_spans):
        raise DocxTableRepairError(
            "Source table border evidence does not match its cell spans"
        )
    normalized_edges = tuple(edges_by_span[span] for span in normalized_spans)

    row_count, column_count = shape
    coverage = [[None] * column_count for _row in range(row_count)]
    for span_index, (row, column, row_span, column_span) in enumerate(
        normalized_spans
    ):
        for covered_row in range(row, row + row_span):
            for covered_column in range(column, column + column_span):
                coverage[covered_row][covered_column] = span_index

    for row in range(row_count):
        for column in range(column_count - 1):
            left_owner = coverage[row][column]
            right_owner = coverage[row][column + 1]
            if left_owner == right_owner:
                continue
            if (
                normalized_edges[left_owner].right
                != normalized_edges[right_owner].left
            ):
                raise DocxTableRepairError(
                    "Source table has conflicting vertical border evidence"
                )
    for row in range(row_count - 1):
        for column in range(column_count):
            top_owner = coverage[row][column]
            bottom_owner = coverage[row + 1][column]
            if top_owner == bottom_owner:
                continue
            if (
                normalized_edges[top_owner].bottom
                != normalized_edges[bottom_owner].top
            ):
                raise DocxTableRepairError(
                    "Source table has conflicting horizontal border evidence"
                )
    return normalized_edges


def _matrix_matches(source, output) -> bool:
    if len(source) != len(output):
        return False
    for source_row, output_row in zip(source, output):
        if len(source_row) != len(output_row):
            return False
        for source_value, output_value in zip(source_row, output_row):
            if (source_value is None) != (output_value is None):
                return False
            if source_value is not None and (
                _normalize_cell_text(source_value)
                != _normalize_cell_text(output_value)
            ):
                return False
    return True


def _contains_forbidden_content(cell) -> bool:
    for element in cell.iter():
        if element is cell:
            continue
        if element.tag == _qname("tbl") or element.tag in _FORBIDDEN_CONTENT_TAGS:
            return True
    return False


def _canonical_cell_style(cell) -> bytes:
    properties = _cell_properties(cell)
    if properties is None:
        return b""
    clone = deepcopy(properties)
    ignored = {
        _qname("gridSpan"),
        _qname("tcBorders"),
        _qname("tcW"),
        _qname("vMerge"),
    }
    for child in list(clone):
        if child.tag == _qname("tcPrChange"):
            raise DocxTableRepairError("Tracked cell-property changes are unsafe")
        if child.tag in ignored:
            clone.remove(child)
    return ElementTree.tostring(clone, encoding="utf-8")


def _border_edge(cell, name: str):
    properties = _cell_properties(cell)
    borders = (
        properties.find(f"./{_qname('tcBorders')}")
        if properties is not None
        else None
    )
    return borders.find(f"./{_qname(name)}") if borders is not None else None


def _canonical_element(element) -> bytes:
    return b"" if element is None else ElementTree.tostring(element, encoding="utf-8")


def _validate_group_style(group) -> None:
    styles = {_canonical_cell_style(item.element) for item in group}
    if len(styles) != 1:
        raise DocxTableRepairError("Split cells have conflicting cell styles")
    for diagonal in ("tl2br", "tr2bl", "insideH", "insideV"):
        if any(_border_edge(item.element, diagonal) is not None for item in group):
            raise DocxTableRepairError("Complex cell borders cannot be merged safely")
    for edge in ("top", "bottom"):
        values = {
            _canonical_element(_border_edge(item.element, edge))
            for item in group
        }
        if len(values) != 1:
            raise DocxTableRepairError("Split cells have conflicting outer borders")


def _group_content_gate(group, expected_text) -> None:
    nonempty = [item for item in group if _normalize_cell_text(item.text)]
    expected = _normalize_cell_text(expected_text)
    if expected:
        if len(nonempty) != 1 or _normalize_cell_text(nonempty[0].text) != expected:
            raise DocxTableRepairError("Split-cell text does not match the PDF cell")
    elif nonempty:
        raise DocxTableRepairError("A blank PDF cell contains Word text")
    if any(_contains_forbidden_content(item.element) for item in group):
        raise DocxTableRepairError("Cell contains an object that cannot be moved safely")
    if len(group) > 1:
        _validate_group_style(group)


def _row_boundary_positions(row, total_width: float) -> tuple[float, ...]:
    desired_total = sum(item.desired_width for item in row)
    if desired_total <= 0:
        raise DocxTableRepairError("Word row widths are invalid")
    scale = total_width / desired_total
    positions = [0.0]
    running = 0.0
    for item in row:
        running += item.desired_width * scale
        positions.append(running)
    positions[-1] = total_width
    return tuple(positions)


def _source_boundary_positions(widths, total_width: float) -> tuple[float, ...]:
    source_total = sum(widths)
    if source_total <= 0:
        raise DocxTableRepairError("Source column widths are invalid")
    scale = total_width / source_total
    positions = [0.0]
    running = 0.0
    for width in widths:
        running += width * scale
        positions.append(running)
    positions[-1] = total_width
    return tuple(positions)


def _paragraph_runs(paragraph):
    runs = []
    properties_seen = False
    allowed_payload = {
        _qname("lastRenderedPageBreak"),
        _qname("t"),
        _qname("tab"),
    }
    for child in paragraph:
        if child.tag == _qname("pPr") and not properties_seen and not runs:
            properties_seen = True
            continue
        if child.tag != _qname("r"):
            raise DocxTableRepairError(
                "Merged-cell paragraph contains unsupported content"
            )
        payload = [item for item in child if item.tag != _qname("rPr")]
        if any(item.tag not in allowed_payload for item in payload):
            raise DocxTableRepairError(
                "Merged-cell run cannot be split without losing content"
            )
        runs.append(child)
    return tuple(runs)


def _unique_paragraph_run_partition(paragraph, expected_texts):
    runs = _paragraph_runs(paragraph)
    run_texts = tuple(_normalize_cell_text(_cell_text(run)) for run in runs)
    content_indices = tuple(
        index for index, value in enumerate(run_texts) if value
    )
    expected = tuple(_normalize_cell_text(value) for value in expected_texts)
    solutions = []

    def search(expected_index, content_index, groups):
        if len(solutions) > 1:
            return
        if expected_index == len(expected):
            if content_index == len(content_indices):
                solutions.append(tuple(groups))
            return
        expected_value = expected[expected_index]
        if not expected_value:
            search(expected_index + 1, content_index, groups + ((),))
            return
        if content_index >= len(content_indices):
            return
        first_run = content_indices[content_index]
        for next_content_index in range(
            content_index + 1,
            len(content_indices) + 1,
        ):
            last_run = content_indices[next_content_index - 1]
            group = runs[first_run : last_run + 1]
            group_text = _normalize_cell_text(
                "".join(_cell_text(run) for run in group)
            )
            if group_text != expected_value:
                continue
            search(
                expected_index + 1,
                next_content_index,
                groups + (tuple(group),),
            )

    search(0, 0, ())
    return solutions[0] if len(solutions) == 1 else None


def _paragraph_from_runs(paragraph, runs):
    result = ElementTree.Element(_qname("p"))
    properties = paragraph.find(f"./{_qname('pPr')}")
    if properties is not None:
        result.append(deepcopy(properties))
    for run in runs:
        clone = deepcopy(run)
        for marker in clone.findall(f"./{_qname('lastRenderedPageBreak')}"):
            clone.remove(marker)
        result.append(clone)
    return result


def _grid_boundary_positions(grid_widths) -> tuple[float, ...]:
    positions = [0.0]
    running = 0.0
    for width in grid_widths:
        running += float(width)
        positions.append(running)
    return tuple(positions)


def _unique_current_grid_partition(
    cell,
    source_segments,
    source_positions,
    grid_widths,
    tolerance,
):
    if len(source_segments) == 1:
        return (cell.end - cell.start,)
    grid_positions = _grid_boundary_positions(grid_widths)
    solutions = []

    def search(segment_index, grid_index, boundaries):
        if len(solutions) > 1:
            return
        if segment_index == len(source_segments) - 1:
            if grid_index < cell.end:
                solutions.append(tuple(boundaries) + (cell.end,))
            return
        remaining = len(source_segments) - segment_index - 1
        boundary = source_positions[source_segments[segment_index].end]
        for candidate in range(
            grid_index + 1,
            cell.end - remaining + 1,
        ):
            if abs(grid_positions[candidate] - boundary) <= tolerance:
                search(
                    segment_index + 1,
                    candidate,
                    boundaries + (candidate,),
                )

    search(0, cell.start, (cell.start,))
    if len(solutions) != 1:
        return None
    boundaries = solutions[0]
    spans = tuple(
        end - start for start, end in zip(boundaries, boundaries[1:])
    )
    if len(spans) != len(source_segments) or any(span < 1 for span in spans):
        return None
    return spans


def _unique_vertical_aggregate_row_plan(
    row,
    cell,
    paragraph,
    source_segments,
    matrix,
    source_positions,
    grid_widths,
    total_width,
):
    row_positions = _row_boundary_positions(row, total_width)
    cell_index = next(
        index for index, item in enumerate(row) if item.element is cell.element
    )
    output_start = row_positions[cell_index]
    output_end = row_positions[cell_index + 1]
    tolerance = max(40.0, total_width * 0.003)
    candidates = []
    for start_index, first in enumerate(source_segments):
        if abs(source_positions[first.start] - output_start) > tolerance:
            continue
        for end_index in range(start_index + 1, len(source_segments) + 1):
            selected = source_segments[start_index:end_index]
            if abs(source_positions[selected[-1].end] - output_end) > tolerance:
                continue
            if any(
                item.continuation or item.row_span != 1
                for item in selected
            ):
                continue
            expected_texts = tuple(
                matrix[item.owner_row][item.owner_column]
                for item in selected
            )
            run_groups = _unique_paragraph_run_partition(
                paragraph,
                expected_texts,
            )
            if run_groups is None:
                continue
            grid_spans = _unique_current_grid_partition(
                cell,
                selected,
                source_positions,
                grid_widths,
                tolerance,
            )
            if grid_spans is None:
                continue
            candidates.append((tuple(selected), run_groups, grid_spans))
    return candidates[0] if len(candidates) == 1 else None


def _replace_cell_with_expanded_segments(
    row_element,
    cell,
    paragraph,
    source_segments,
    run_groups,
    grid_spans,
    grid_widths,
):
    insertion_index = row_element.index(cell.element)
    template = deepcopy(cell.element)
    row_element.remove(cell.element)
    grid_index = cell.start
    for offset, (segment, runs, grid_span) in enumerate(
        zip(source_segments, run_groups, grid_spans)
    ):
        expanded = deepcopy(template)
        properties = _cell_properties(expanded, create=True)
        for child in list(expanded):
            if child is not properties:
                expanded.remove(child)
        expanded.append(_paragraph_from_runs(paragraph, runs))
        _set_grid_span(expanded, grid_span)
        _set_cell_width(
            expanded,
            sum(grid_widths[grid_index : grid_index + grid_span]),
        )
        _set_vertical_merge(expanded, segment)
        row_element.insert(insertion_index + offset, expanded)
        grid_index += grid_span


def _expand_uniquely_proven_vertical_aggregates(
    table,
    source_segments,
    matrix,
    source_positions,
    grid_widths,
    total_width,
) -> bool:
    output_rows = _output_rows(table, grid_widths)
    row_elements = table.findall(f"./{_qname('tr')}")
    lookups = tuple(
        {(cell.start, cell.end): cell for cell in row}
        for row in output_rows
    )
    plans = []
    for row_index, row in enumerate(output_rows):
        for cell in row:
            if _vertical_merge_value(cell.element) != "restart":
                continue
            chain = [cell]
            next_row = row_index + 1
            while next_row < len(output_rows):
                continuation = lookups[next_row].get((cell.start, cell.end))
                if (
                    continuation is None
                    or _vertical_merge_value(continuation.element) != "continue"
                ):
                    break
                chain.append(continuation)
                next_row += 1
            if len(chain) < 2:
                continue
            if any(_contains_forbidden_content(item.element) for item in chain):
                continue
            if len({_canonical_cell_style(item.element) for item in chain}) != 1:
                continue
            paragraphs = tuple(
                child for child in cell.element if child.tag == _qname("p")
            )
            non_properties = tuple(
                child
                for child in cell.element
                if child.tag != _qname("tcPr")
            )
            if len(paragraphs) != len(chain) or paragraphs != non_properties:
                continue
            if any(_normalize_cell_text(item.text) for item in chain[1:]):
                continue
            chain_plans = []
            for offset, (chain_cell, paragraph) in enumerate(
                zip(chain, paragraphs)
            ):
                source_row_index = row_index + offset
                plan = _unique_vertical_aggregate_row_plan(
                    output_rows[source_row_index],
                    chain_cell,
                    paragraph,
                    source_segments[source_row_index],
                    matrix,
                    source_positions,
                    grid_widths,
                    total_width,
                )
                if plan is None:
                    chain_plans = []
                    break
                chain_plans.append(
                    (
                        row_elements[source_row_index],
                        chain_cell,
                        paragraph,
                        *plan,
                    )
                )
            if chain_plans:
                plans.extend(chain_plans)
    for plan in sorted(plans, key=lambda item: item[1].start, reverse=True):
        _replace_cell_with_expanded_segments(*plan, grid_widths)
    return bool(plans)


def _unique_row_plan(row, segments, matrix, source_positions, total_width):
    output_positions = _row_boundary_positions(row, total_width)
    tolerance = max(40.0, total_width * 0.003)
    solutions = []

    def search(segment_index, cell_index, groups):
        if len(solutions) > 1:
            return
        if segment_index == len(segments):
            if cell_index == len(row):
                solutions.append(tuple(groups))
            return
        segment = segments[segment_index]
        expected_start = source_positions[segment.start]
        if abs(output_positions[cell_index] - expected_start) > tolerance:
            return
        expected_end = source_positions[segment.end]
        expected_text = (
            None
            if segment.continuation
            else matrix[segment.owner_row][segment.owner_column]
        )
        for end_index in range(cell_index + 1, len(row) + 1):
            if abs(output_positions[end_index] - expected_end) > tolerance:
                continue
            group = row[cell_index:end_index]
            try:
                _group_content_gate(group, expected_text)
            except DocxTableRepairError:
                continue
            search(
                segment_index + 1,
                end_index,
                groups + ((segment, group),),
            )

    search(0, 0, ())
    if len(solutions) != 1:
        raise DocxTableRepairError(
            "Word cells do not have one unique mapping to the PDF grid"
        )
    return solutions[0]


def _scaled_integer_widths(widths, total: int) -> tuple[int, ...]:
    if total < len(widths):
        raise DocxTableRepairError("Word table is too narrow for the source grid")
    source_total = sum(widths)
    raw = [width / source_total * total for width in widths]
    values = [max(1, int(value)) for value in raw]
    difference = total - sum(values)
    if difference > 0:
        order = sorted(
            range(len(values)),
            key=lambda index: raw[index] - int(raw[index]),
            reverse=True,
        )
        for index in range(difference):
            values[order[index % len(order)]] += 1
    elif difference < 0:
        order = sorted(
            range(len(values)),
            key=lambda index: (values[index], raw[index] - int(raw[index])),
            reverse=True,
        )
        remaining = -difference
        for index in order:
            removable = min(remaining, values[index] - 1)
            values[index] -= removable
            remaining -= removable
            if not remaining:
                break
        if remaining:
            raise DocxTableRepairError("Cannot scale the source grid safely")
    return tuple(values)


def _replace_grid(table, widths) -> None:
    grid = table.find(f"./{_qname('tblGrid')}")
    if grid is None:
        raise DocxTableRepairError("Word table has no tblGrid definition")
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        column = ElementTree.Element(_qname("gridCol"))
        column.set(_qname("w"), str(int(width)))
        grid.append(column)


def _set_grid_span(cell, span: int) -> None:
    properties = _cell_properties(cell, create=True)
    element = properties.find(f"./{_qname('gridSpan')}")
    if span <= 1:
        if element is not None:
            properties.remove(element)
        return
    if element is None:
        element = ElementTree.Element(_qname("gridSpan"))
        properties.append(element)
    element.set(_qname("val"), str(int(span)))


def _set_cell_width(cell, width: int) -> None:
    properties = _cell_properties(cell, create=True)
    element = properties.find(f"./{_qname('tcW')}")
    if element is None:
        element = ElementTree.Element(_qname("tcW"))
        properties.insert(0, element)
    element.set(_qname("type"), "dxa")
    element.set(_qname("w"), str(int(width)))


def _set_vertical_merge(cell, segment: _SourceSegment) -> None:
    properties = _cell_properties(cell, create=True)
    element = properties.find(f"./{_qname('vMerge')}")
    if segment.row_span <= 1:
        if element is not None:
            properties.remove(element)
        return
    if element is None:
        element = ElementTree.Element(_qname("vMerge"))
        properties.append(element)
    if segment.continuation:
        element.attrib.pop(_qname("val"), None)
    else:
        element.set(_qname("val"), "restart")


def _copy_border_edge(target, source, name: str) -> None:
    for existing in list(target.findall(f"./{_qname(name)}")):
        target.remove(existing)
    if source is not None:
        target.append(deepcopy(source))


def _cell_border_elements(cell, name: str):
    properties = _cell_properties(cell)
    borders = (
        properties.find(f"./{_qname('tcBorders')}")
        if properties is not None
        else None
    )
    if borders is None:
        return ()
    return tuple(
        element
        for alias in _CELL_BORDER_EDGE_ALIASES[name]
        for element in borders.findall(f"./{_qname(alias)}")
    )


def _border_value(border) -> str:
    return str(border.get(_qname("val"), "")).strip().lower()


def _border_is_visible(border) -> bool:
    value = _border_value(border)
    return bool(value and value not in {"nil", "none"})


def _border_style_signature(border):
    return (
        tuple(sorted((str(name), str(value)) for name, value in border.attrib.items())),
        border.text or "",
        tuple(_canonical_element(child) for child in border),
    )


def _unique_visible_cell_border(cell, name: str):
    visible = tuple(
        border
        for border in _cell_border_elements(cell, name)
        if _border_is_visible(border)
    )
    signatures = {_border_style_signature(border) for border in visible}
    if len(signatures) > 1:
        raise DocxTableRepairError("Word cell has conflicting border styles")
    return visible[0] if visible else None


def _validate_simple_cell_borders(table) -> None:
    allowed = {
        alias
        for aliases in _CELL_BORDER_EDGE_ALIASES.values()
        for alias in aliases
    }
    for borders in table.findall(f".//{_qname('tcBorders')}"):
        for border in borders:
            local_name = ElementTree.QName(border).localname
            if (
                local_name in _COMPLEX_CELL_BORDER_NAMES
                or local_name not in allowed
            ):
                raise DocxTableRepairError(
                    "Complex Word cell borders cannot be repaired safely"
                )
        cell = borders.getparent()
        while cell is not None and cell.tag != _qname("tc"):
            cell = cell.getparent()
        if cell is not None:
            for name in _CELL_BORDER_EDGE_NAMES:
                _unique_visible_cell_border(cell, name)


def _visible_borders_compatible(first, second) -> bool:
    return bool(
        first is None
        or second is None
        or _border_style_signature(first) == _border_style_signature(second)
    )


def _validate_shared_border_styles(output_rows, column_count: int) -> None:
    unit_rows = []
    for row in output_rows:
        units = [None] * column_count
        for cell in row:
            for column in range(cell.start, cell.end):
                units[column] = cell.element
        if any(cell is None for cell in units):
            raise DocxTableRepairError("Word table border grid is incomplete")
        unit_rows.append(tuple(units))

    for units in unit_rows:
        for column in range(column_count - 1):
            first = units[column]
            second = units[column + 1]
            if first is second:
                continue
            if not _visible_borders_compatible(
                _unique_visible_cell_border(first, "right"),
                _unique_visible_cell_border(second, "left"),
            ):
                raise DocxTableRepairError(
                    "Adjacent Word cells have conflicting border styles"
                )
    for top_units, bottom_units in zip(unit_rows, unit_rows[1:]):
        for column in range(column_count):
            first = top_units[column]
            second = bottom_units[column]
            if first is second:
                continue
            if not _visible_borders_compatible(
                _unique_visible_cell_border(first, "bottom"),
                _unique_visible_cell_border(second, "top"),
            ):
                raise DocxTableRepairError(
                    "Adjacent Word cells have conflicting border styles"
                )


def _table_common_visible_border(table):
    candidates = []
    for cell in table.findall(f".//{_qname('tc')}"):
        for name in _CELL_BORDER_EDGE_NAMES:
            border = _unique_visible_cell_border(cell, name)
            if border is not None:
                candidates.append(border)
    table_borders = table.find(
        f"./{_qname('tblPr')}/{_qname('tblBorders')}"
    )
    if table_borders is not None:
        candidates.extend(
            border for border in table_borders if _border_is_visible(border)
        )
    if not candidates:
        return None

    counts = {}
    first_by_signature = {}
    order = []
    for border in candidates:
        signature = _border_style_signature(border)
        if signature not in counts:
            counts[signature] = 0
            first_by_signature[signature] = border
            order.append(signature)
        counts[signature] += 1
    winner = max(order, key=lambda signature: counts[signature])
    return first_by_signature[winner]


def _physical_cells_for_source_spans(table, grid_widths, shape, spans):
    row_count, column_count = shape
    if len(grid_widths) != column_count:
        raise DocxTableRepairError("Word table grid does not match the PDF")
    output_rows = _output_rows(table, grid_widths)
    if len(output_rows) != row_count:
        raise DocxTableRepairError("Word row count does not match the PDF")
    lookups = tuple(
        {(cell.start, cell.end): cell.element for cell in row}
        for row in output_rows
    )
    chains = []
    for row, column, row_span, column_span in spans:
        key = (column, column + column_span)
        chain = []
        for offset in range(row_span):
            cell = lookups[row + offset].get(key)
            if cell is None:
                raise DocxTableRepairError(
                    "Word cells do not match the PDF border grid"
                )
            merge_value = _vertical_merge_value(cell)
            if row_span == 1:
                if merge_value:
                    raise DocxTableRepairError(
                        "Word cell has an unexpected vertical merge"
                    )
            elif offset == 0:
                if merge_value != "restart":
                    raise DocxTableRepairError(
                        "Word vertical merge does not match the PDF"
                    )
            elif merge_value != "continue":
                raise DocxTableRepairError(
                    "Word vertical merge does not match the PDF"
                )
            chain.append(cell)
        chains.append(tuple(chain))
    return tuple(output_rows), tuple(chains)


def _logical_edge_template(chain, name: str):
    selected_cells = (
        (chain[0],)
        if name == "top"
        else (chain[-1],)
        if name == "bottom"
        else chain
    )
    visible = tuple(
        border
        for cell in selected_cells
        if (border := _unique_visible_cell_border(cell, name)) is not None
    )
    signatures = {_border_style_signature(border) for border in visible}
    if len(signatures) > 1:
        raise DocxTableRepairError(
            "A merged Word cell has conflicting border styles"
        )
    return visible[0] if visible else None


def _build_table_border_plan(table, grid_widths, shape, spans, border_edges):
    _validate_simple_cell_borders(table)
    output_rows, chains = _physical_cells_for_source_spans(
        table,
        grid_widths,
        shape,
        spans,
    )
    _validate_shared_border_styles(output_rows, len(grid_widths))
    common_template = _table_common_visible_border(table)
    plan = []
    for chain, evidence in zip(chains, border_edges):
        logical_templates = {}
        for name in _CELL_BORDER_EDGE_NAMES:
            if not getattr(evidence, name):
                logical_templates[name] = None
                continue
            template = _logical_edge_template(chain, name)
            logical_templates[name] = (
                common_template if template is None else template
            )
        for offset, cell in enumerate(chain):
            desired = {
                "top": evidence.top if offset == 0 else False,
                "left": evidence.left,
                "bottom": evidence.bottom if offset == len(chain) - 1 else False,
                "right": evidence.right,
            }
            templates = {}
            for name in _CELL_BORDER_EDGE_NAMES:
                if not desired[name]:
                    templates[name] = None
                    continue
                template = _unique_visible_cell_border(cell, name)
                templates[name] = (
                    logical_templates[name] if template is None else template
                )
            plan.append((cell, desired, templates))
    return tuple(plan)


def _make_border_element(name: str, template=None, *, visible: bool):
    if not visible:
        border = ElementTree.Element(_qname(name))
        border.set(_qname("val"), "nil")
        return border
    if template is not None:
        border = deepcopy(template)
        border.tag = _qname(name)
        return border
    border = ElementTree.Element(_qname(name))
    border.set(_qname("val"), "single")
    border.set(_qname("sz"), "4")
    border.set(_qname("space"), "0")
    border.set(_qname("color"), "000000")
    return border


def _insert_cell_border(borders, border, name: str) -> None:
    order = {edge: index for index, edge in enumerate(_CELL_BORDER_EDGE_NAMES)}
    target_order = order[name]
    for index, child in enumerate(borders):
        local_name = ElementTree.QName(child).localname
        canonical_name = (
            "left"
            if local_name == "start"
            else "right"
            if local_name == "end"
            else local_name
        )
        if order.get(canonical_name, len(order)) > target_order:
            borders.insert(index, border)
            return
    borders.append(border)


def _ensure_cell_border_edge(cell, name: str, desired: bool, template=None) -> bool:
    elements = _cell_border_elements(cell, name)
    visible = tuple(border for border in elements if _border_is_visible(border))
    if desired:
        if visible and len(visible) == len(elements):
            return False
    elif (
        elements
        and not visible
        and all(_border_value(border) == "nil" for border in elements)
    ):
        return False

    properties = _cell_properties(cell, create=True)
    borders = properties.find(f"./{_qname('tcBorders')}")
    if borders is None:
        borders = ElementTree.Element(_qname("tcBorders"))
        properties.append(borders)
    aliases = {_qname(alias) for alias in _CELL_BORDER_EDGE_ALIASES[name]}
    for existing in list(borders):
        if existing.tag in aliases:
            borders.remove(existing)
    _insert_cell_border(
        borders,
        _make_border_element(name, template, visible=desired),
        name,
    )
    return True


def _apply_table_border_plan(plan) -> bool:
    changed = False
    for cell, desired, templates in plan:
        for name in _CELL_BORDER_EDGE_NAMES:
            changed = (
                _ensure_cell_border_edge(
                    cell,
                    name,
                    desired[name],
                    templates[name],
                )
                or changed
            )
    return changed


def _verify_table_border_evidence(table, grid_widths, shape, spans, border_edges):
    output_rows, chains = _physical_cells_for_source_spans(
        table,
        grid_widths,
        shape,
        spans,
    )
    for chain, evidence in zip(chains, border_edges):
        for offset, cell in enumerate(chain):
            expected = {
                "top": evidence.top if offset == 0 else False,
                "left": evidence.left,
                "bottom": evidence.bottom if offset == len(chain) - 1 else False,
                "right": evidence.right,
            }
            for name, should_be_visible in expected.items():
                elements = _cell_border_elements(cell, name)
                is_visible = any(_border_is_visible(item) for item in elements)
                has_explicit_nil = bool(
                    elements
                    and not is_visible
                    and any(_border_value(item) == "nil" for item in elements)
                )
                if is_visible != should_be_visible or (
                    not should_be_visible and not has_explicit_nil
                ):
                    raise DocxTableRepairError(
                        "Word cell borders do not match the PDF"
                    )
    _validate_shared_border_styles(output_rows, len(grid_widths))


def _merge_outer_borders(survivor, group) -> None:
    properties = _cell_properties(survivor, create=True)
    borders = properties.find(f"./{_qname('tcBorders')}")
    edges = {
        "top": _border_edge(group[0].element, "top"),
        "bottom": _border_edge(group[0].element, "bottom"),
        "left": _border_edge(group[0].element, "left"),
        "start": _border_edge(group[0].element, "start"),
        "right": _border_edge(group[-1].element, "right"),
        "end": _border_edge(group[-1].element, "end"),
    }
    if borders is None and any(value is not None for value in edges.values()):
        borders = ElementTree.Element(_qname("tcBorders"))
        properties.append(borders)
    if borders is None:
        return
    for name in ("insideH", "insideV", "tl2br", "tr2bl"):
        for existing in list(borders.findall(f"./{_qname(name)}")):
            borders.remove(existing)
    for name, source in edges.items():
        _copy_border_edge(borders, source, name)
    if not list(borders):
        properties.remove(borders)


def _move_blocks_to_survivor(survivor, donor) -> None:
    survivor_properties = _cell_properties(survivor, create=True)
    if donor is survivor:
        return
    donor_properties = _cell_properties(donor)
    survivor_blocks = [
        child for child in list(survivor) if child is not survivor_properties
    ]
    for child in survivor_blocks:
        survivor.remove(child)
    donor_blocks = [
        child for child in list(donor) if child is not donor_properties
    ]
    for child in donor_blocks:
        _remove_empty_placeholder_runs(child)
        donor.remove(child)
        survivor.append(child)
    if not any(child.tag == _qname("p") for child in survivor):
        survivor.append(ElementTree.Element(_qname("p")))


def _remove_empty_placeholder_runs(parent) -> None:
    for child in list(parent):
        if child.tag != _qname("r"):
            _remove_empty_placeholder_runs(child)
            continue
        meaningful = False
        for descendant in child:
            if descendant.tag == _qname("rPr"):
                continue
            if descendant.tag == _qname("t") and not (descendant.text or ""):
                continue
            meaningful = True
            break
        if not meaningful:
            parent.remove(child)


def _apply_table_plan(table, row_plans, new_widths) -> None:
    rows = table.findall(f"./{_qname('tr')}")
    for row, row_plan in zip(rows, row_plans):
        for segment, group in row_plan:
            survivor = group[0].element
            nonempty = [item for item in group if _normalize_cell_text(item.text)]
            donor = nonempty[0].element if nonempty else survivor
            _move_blocks_to_survivor(survivor, donor)
            _set_grid_span(survivor, segment.end - segment.start)
            _set_cell_width(
                survivor,
                sum(new_widths[segment.start : segment.end]),
            )
            _set_vertical_merge(survivor, segment)
            if len(group) > 1:
                _merge_outer_borders(survivor, group)
            for item in group[1:]:
                row.remove(item.element)
    _replace_grid(table, new_widths)


def _table_properties(table, *, create=False):
    properties = table.find(f"./{_qname('tblPr')}")
    if properties is None and create:
        properties = ElementTree.Element(_qname("tblPr"))
        table.insert(0, properties)
    return properties


def _insert_before_first(parent, element, following_names) -> None:
    following_tags = {_qname(name) for name in following_names}
    for index, child in enumerate(parent):
        if child.tag in following_tags:
            parent.insert(index, element)
            return
    parent.append(element)


def _cell_has_complete_outer_borders(cell) -> bool:
    properties = _cell_properties(cell)
    borders = (
        properties.find(f"./{_qname('tcBorders')}")
        if properties is not None
        else None
    )
    if borders is None:
        return False

    def visible_edge(*names):
        for name in names:
            edge = borders.find(f"./{_qname(name)}")
            if edge is None:
                continue
            value = str(edge.get(_qname("val"), "")).strip().lower()
            if value and value not in {"nil", "none"}:
                return True
        return False

    return (
        visible_edge("top")
        and visible_edge("bottom")
        and visible_edge("left", "start")
        and visible_edge("right", "end")
    )


def _run_border_signature(border):
    return (
        border.tag,
        tuple(sorted((str(name), str(value)) for name, value in border.attrib.items())),
        border.text or "",
        tuple(_canonical_element(child) for child in border),
    )


def _remove_systematic_run_borders(table, minimum_repetition=4) -> int:
    parent_map = {
        child: parent
        for parent in table.iter()
        for child in parent
    }
    groups = {}
    for properties in table.findall(f".//{_qname('rPr')}"):
        for border in properties.findall(f"./{_qname('bdr')}"):
            ancestor = parent_map.get(properties)
            while ancestor is not None and ancestor.tag != _qname("tc"):
                ancestor = parent_map.get(ancestor)
            if ancestor is None:
                continue
            signature = _run_border_signature(border)
            groups.setdefault(signature, []).append(
                (properties, border, ancestor)
            )

    removed = 0
    for entries in groups.values():
        distinct_cells = {id(cell) for _properties, _border, cell in entries}
        if (
            len(entries) < minimum_repetition
            or len(distinct_cells) < minimum_repetition
            or any(
                not _cell_has_complete_outer_borders(cell)
                for _properties, _border, cell in entries
            )
        ):
            continue
        for properties, border, _cell in entries:
            properties.remove(border)
            removed += 1
    return removed


def _positive_table_row_heights(table):
    rows = table.findall(f"./{_qname('tr')}")
    if not rows:
        return None
    result = []
    for row in rows:
        properties = row.find(f"./{_qname('trPr')}")
        heights = (
            properties.findall(f"./{_qname('trHeight')}")
            if properties is not None
            else ()
        )
        if not heights:
            return None
        values = tuple(
            _integer_attribute(height, "val", 0)
            for height in heights
        )
        if any(value is None or value <= 0 for value in values):
            return None
        result.append((tuple(heights), max(int(value) for value in values)))
    return tuple(result)


def _set_exact_row_heights(table):
    row_heights = _positive_table_row_heights(table)
    if row_heights is None:
        return None
    for height_elements, _value in row_heights:
        for height in height_elements:
            height.set(_qname("hRule"), "exact")
    return tuple(value for _height_elements, value in row_heights)


def _set_table_vertical_cell_margins_zero(table) -> None:
    properties = _table_properties(table, create=True)
    margins = properties.find(f"./{_qname('tblCellMar')}")
    if margins is None:
        margins = ElementTree.Element(_qname("tblCellMar"))
        _insert_before_first(
            properties,
            margins,
            ("tblLook", "tblCaption", "tblDescription", "tblPrChange"),
        )

    top = margins.find(f"./{_qname('top')}")
    if top is None:
        top = ElementTree.Element(_qname("top"))
        margins.insert(0, top)
    bottom = margins.find(f"./{_qname('bottom')}")
    if bottom is None:
        bottom = ElementTree.Element(_qname("bottom"))
        _insert_before_first(margins, bottom, ("right", "end"))
    for edge in (top, bottom):
        edge.set(_qname("w"), "0")
        edge.set(_qname("type"), "dxa")


def _set_table_indent_zero(table) -> None:
    properties = _table_properties(table, create=True)
    indent = properties.find(f"./{_qname('tblInd')}")
    if indent is None:
        indent = ElementTree.Element(_qname("tblInd"))
        _insert_before_first(
            properties,
            indent,
            (
                "tblBorders",
                "shd",
                "tblLayout",
                "tblCellMar",
                "tblLook",
                "tblCaption",
                "tblDescription",
                "tblPrChange",
            ),
        )
    indent.set(_qname("w"), "0")
    indent.set(_qname("type"), "dxa")


def _repair_table_layout_artifacts(table):
    row_heights = _set_exact_row_heights(table)
    _set_table_vertical_cell_margins_zero(table)
    return row_heights


def _tighten_single_full_page_table(
    root,
    repaired_tables,
    row_heights_by_table,
) -> bool:
    body = root.find(f"./{_qname('body')}")
    if body is None:
        return False
    direct_tables = tuple(body.findall(f"./{_qname('tbl')}"))
    if len(direct_tables) != 1:
        return False
    table = direct_tables[0]
    if not any(table is candidate for candidate in repaired_tables):
        return False
    row_heights = row_heights_by_table.get(id(table))
    if row_heights is None:
        return False

    section = body.find(f"./{_qname('sectPr')}")
    page_size = (
        section.find(f"./{_qname('pgSz')}")
        if section is not None
        else None
    )
    page_width = _integer_attribute(page_size, "w", 0)
    page_height = _integer_attribute(page_size, "h", 0)
    if not page_width or not page_height:
        return False
    table_width = sum(_table_grid_widths(table))
    total_row_height = sum(row_heights)
    if (
        table_width <= page_width * 0.8
        or table_width > page_width
        or total_row_height <= page_height * 0.6
    ):
        return False

    page_margins = section.find(f"./{_qname('pgMar')}")
    if page_margins is None:
        page_margins = ElementTree.Element(_qname("pgMar"))
        _insert_before_first(
            section,
            page_margins,
            (
                "paperSrc",
                "pgBorders",
                "lnNumType",
                "pgNumType",
                "cols",
                "formProt",
                "vAlign",
                "noEndnote",
                "titlePg",
                "textDirection",
                "bidi",
                "rtlGutter",
                "docGrid",
                "printerSettings",
                "sectPrChange",
            ),
        )
    horizontal_margin = max(0, int(round((page_width - table_width) / 2)))
    for name, value in (
        ("top", 160),
        ("bottom", 160),
        ("left", horizontal_margin),
        ("right", horizontal_margin),
    ):
        page_margins.set(_qname(name), str(value))
    _set_table_indent_zero(table)
    return True


def _validate_markup_compatibility_namespaces(root) -> None:
    for element in root.iter():
        namespaces = element.nsmap
        for attribute_name in _MARKUP_PREFIX_ATTRIBUTES:
            value = element.get(
                f"{{{_MARKUP_COMPATIBILITY_NAMESPACE}}}{attribute_name}"
            )
            if not value:
                continue
            for token in str(value).split():
                prefix = token.split(":", 1)[0]
                if prefix not in namespaces:
                    raise DocxTableRepairError(
                        "Word markup-compatibility namespace is missing: "
                        f"{prefix}"
                    )


def _write_repaired_archive(path: Path, document_xml: bytes) -> None:
    temporary = path.with_name(f".{path.name}.{uuid4().hex}.repair")
    try:
        with ZipFile(path, "r") as source, ZipFile(temporary, "w") as target:
            for info in source.infolist():
                data = (
                    document_xml
                    if info.filename == "word/document.xml"
                    else source.read(info.filename)
                )
                target.writestr(info, data)
        with ZipFile(temporary, "r") as archive:
            repaired_root = ElementTree.fromstring(
                archive.read("word/document.xml")
            )
            _validate_markup_compatibility_namespaces(repaired_root)
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def repair_docx_table_topology(
    path,
    *,
    table_shapes,
    table_cell_matrices,
    table_cell_spans,
    table_column_widths,
    table_cell_border_edges=None,
) -> bool:
    """Repair Word table topology and any explicitly proven PDF cell borders."""

    docx_path = Path(path)
    shapes = tuple(table_shapes or ())
    matrices = tuple(table_cell_matrices or ())
    spans = tuple(table_cell_spans or ())
    widths = tuple(table_column_widths or ())
    border_tables = None
    if table_cell_border_edges is not None:
        try:
            border_tables = tuple(
                tuple(table_edges) for table_edges in table_cell_border_edges
            )
        except TypeError as exc:
            raise DocxTableRepairError(
                "Source table border evidence is malformed"
            ) from exc
    lengths = {len(shapes), len(matrices), len(spans), len(widths)}
    if border_tables is not None:
        lengths.add(len(border_tables))
    if len(lengths) != 1 or not shapes:
        raise DocxTableRepairError("Source table evidence is incomplete")
    source_values = tuple(zip(shapes, matrices, spans, widths))

    try:
        with ZipFile(docx_path, "r") as archive:
            document_xml = archive.read("word/document.xml")
    except (BadZipFile, KeyError, OSError) as exc:
        raise DocxTableRepairError(f"Cannot open Word document: {exc}") from exc

    try:
        root = ElementTree.fromstring(document_xml)
    except ElementTree.ParseError as exc:
        raise DocxTableRepairError("Word document XML is damaged") from exc
    _validate_markup_compatibility_namespaces(root)
    tables = tuple(root.findall(f".//{_qname('tbl')}"))
    if len(tables) != len(source_values):
        raise DocxTableRepairError("Word table count does not match the PDF")

    changed = False
    pending_repairs = []
    repaired_flags = []
    coerced_sources = []
    normalized_border_tables = []
    for table_index, (table, source_value) in enumerate(
        zip(tables, source_values)
    ):
        shape, matrix, source_spans, source_widths, segments = (
            _coerce_source_table(*source_value)
        )
        coerced_sources.append(
            (shape, matrix, source_spans, source_widths, segments)
        )
        normalized_border_tables.append(
            None
            if border_tables is None
            else _coerce_source_border_edges(
                shape,
                source_value[2],
                source_spans,
                border_tables[table_index],
            )
        )
        grid_widths = _table_grid_widths(table)
        output_shape, output_matrix, output_spans = _existing_signature(
            table,
            grid_widths,
        )
        if (
            output_shape == shape
            and output_spans == source_spans
            and _matrix_matches(matrix, output_matrix)
        ):
            pending_repairs.append(None)
            repaired_flags.append(False)
            continue

        output_rows = _output_rows(table, grid_widths)
        if len(segments) != len(output_rows):
            raise DocxTableRepairError("Word row count does not match the PDF")
        _remove_systematic_run_borders(table)
        total_width = float(sum(grid_widths))
        source_positions = _source_boundary_positions(
            source_widths,
            total_width,
        )
        expanded = _expand_uniquely_proven_vertical_aggregates(
            table,
            segments,
            matrix,
            source_positions,
            grid_widths,
            total_width,
        )
        if expanded:
            changed = True
            output_shape, output_matrix, output_spans = _existing_signature(
                table,
                grid_widths,
            )
            if (
                output_shape == shape
                and output_spans == source_spans
                and _matrix_matches(matrix, output_matrix)
            ):
                pending_repairs.append(None)
                repaired_flags.append(True)
                continue
        output_rows = _output_rows(table, grid_widths)
        row_plans = tuple(
            _unique_row_plan(
                output_row,
                source_row,
                matrix,
                source_positions,
                total_width,
            )
            for output_row, source_row in zip(output_rows, segments)
        )
        new_widths = _scaled_integer_widths(source_widths, int(total_width))
        pending_repairs.append((row_plans, new_widths))
        repaired_flags.append(True)
        changed = True

    for table, repair in zip(tables, pending_repairs):
        if repair is not None:
            _apply_table_plan(table, *repair)

    if border_tables is not None:
        border_plans = []
        for table, source, border_edges in zip(
            tables,
            coerced_sources,
            normalized_border_tables,
        ):
            shape, _matrix, source_spans, _source_widths, _segments = source
            grid_widths = _table_grid_widths(table)
            border_plans.append(
                _build_table_border_plan(
                    table,
                    grid_widths,
                    shape,
                    source_spans,
                    border_edges,
                )
            )
        for plan in border_plans:
            changed = _apply_table_border_plan(plan) or changed
        for table, source, border_edges in zip(
            tables,
            coerced_sources,
            normalized_border_tables,
        ):
            shape, _matrix, source_spans, _source_widths, _segments = source
            _verify_table_border_evidence(
                table,
                _table_grid_widths(table),
                shape,
                source_spans,
                border_edges,
            )

    if not changed:
        return False

    repaired_tables = tuple(
        table
        for table, repaired in zip(tables, repaired_flags)
        if repaired
    )
    row_heights_by_table = {
        id(table): _repair_table_layout_artifacts(table)
        for table in repaired_tables
    }
    _tighten_single_full_page_table(
        root,
        repaired_tables,
        row_heights_by_table,
    )

    serialized = ElementTree.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
    )
    _write_repaired_archive(docx_path, serialized)
    return True
