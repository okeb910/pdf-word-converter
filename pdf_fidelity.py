"""Pure PDF fidelity-risk analysis used before choosing a conversion engine."""

from __future__ import annotations

from dataclasses import dataclass
from os import PathLike
from pathlib import Path
import re
import unicodedata
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile

try:
    import pymupdf
except ImportError:  # PyMuPDF releases before the module rename.
    import fitz as pymupdf

if not hasattr(pymupdf, "open"):
    raise ImportError("PyMuPDF 模块缺少 open() API")


_CHECKBOX_SYMBOLS = frozenset("☐☑☒□✓✔√")
_SYMBOL_FONT_NAMES = ("wingdings", "webdings", "symbol")
_DETAILED_TABLE_ANALYSIS_PAGE_LIMIT = 20
_WORD_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DRAWING_NAMESPACE = (
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
)


class PdfFidelityAnalysisCancelled(RuntimeError):
    """Raised when a caller requests a cooperative fidelity-analysis stop."""


def _raise_if_pdf_analysis_cancelled(cancel_requested) -> None:
    if cancel_requested is not None and bool(cancel_requested()):
        raise PdfFidelityAnalysisCancelled("PDF fidelity analysis was cancelled")

@dataclass(frozen=True, slots=True)
class PdfFidelityRisk:
    """Immutable summary of fidelity-sensitive content found in a PDF."""

    path: Path
    page_count: int
    analyzed_pages: int
    table_count: int
    table_text_character_count: int
    table_shapes: tuple[tuple[int, int], ...]
    table_cell_counts: tuple[int, ...]
    table_texts: tuple[str, ...]
    table_cell_matrices: tuple[tuple[tuple[str | None, ...], ...], ...]
    table_text_extraction_failure_count: int
    selectable_text_character_count: int
    editable_table_candidate: bool
    table_layout_suspected: bool
    tagged_table_structure_present: bool
    widget_count: int
    checkbox_symbol_count: int
    vector_mark_count: int
    symbol_font_run_count: int
    vector_path_count: int
    is_complex: bool
    reasons: tuple[str, ...]
    table_full_width_span_rows: tuple[tuple[bool, ...], ...] = ()
    table_cell_spans: tuple[
        tuple[tuple[int, int, int, int], ...], ...
    ] = ()
    table_column_widths: tuple[tuple[float, ...], ...] = ()
    table_cell_border_edges: tuple[
        tuple[tuple[bool, bool, bool, bool], ...], ...
    ] = ()
    unconfirmed_table_region_count: int = 0
    selectable_text: str = ""
    unverifiable_text_character_count: int = 0
    table_analysis_limited: bool = False

    @property
    def table_analysis_uncertain(self) -> bool:
        return bool(self.table_count and self.table_text_extraction_failure_count)

    @property
    def has_selectable_text(self) -> bool:
        return bool(self.selectable_text_character_count)


@dataclass(frozen=True, slots=True)
class EditableDocxTableSummary:
    """Structural evidence that a DOCX contains editable Word table content."""

    table_count: int
    row_count: int
    cell_count: int
    document_protected: bool
    locked_content_control_table_count: int
    table_text_character_count: int
    table_texts: tuple[str, ...]
    table_cell_matrices: tuple[tuple[tuple[str | None, ...], ...], ...]
    table_shapes: tuple[tuple[int, int], ...]
    table_cell_counts: tuple[int, ...]
    drawing_count: int
    large_page_drawing_count: int
    invalid_layout_table_count: int = 0
    table_cell_spans: tuple[
        tuple[tuple[int, int, int, int], ...], ...
    ] = ()
    table_grid_widths: tuple[tuple[int, ...], ...] = ()
    table_raw_cell_counts: tuple[int, ...] = ()
    table_cell_border_edges: tuple[
        tuple[tuple[bool, bool, bool, bool], ...], ...
    ] = ()
    visible_text_character_count: int = 0
    visible_text: str = ""

    @property
    def has_editable_table_structure(self) -> bool:
        return bool(
            self.table_count
            and self.row_count
            and self.cell_count
            and not self.document_protected
            and not self.locked_content_control_table_count
            and not self.invalid_layout_table_count
        )

    @property
    def has_editable_table(self) -> bool:
        return bool(
            self.has_editable_table_structure
            and self.table_text_character_count
        )


def _word_property_is_enabled(element) -> bool:
    if element is None:
        return False
    value = element.get(f"{{{_WORD_NAMESPACE}}}val", "1")
    return str(value).strip().lower() not in {"0", "false", "off", "none"}


def _word_run_properties_hide_text(properties, namespaces) -> bool:
    if properties is None:
        return False
    return any(
        _word_property_is_enabled(properties.find(f"./w:{name}", namespaces))
        for name in ("vanish", "webHidden", "specVanish")
    )


@dataclass(frozen=True, slots=True)
class _WordStyleVisibility:
    hidden_style_ids: frozenset[str]
    near_white_style_ids: frozenset[str]
    explicit_color_style_ids: frozenset[str]
    invalid_conditional_table_style_ids: frozenset[str]
    default_hidden: bool
    default_near_white: bool
    default_table_style_id: str


def _word_style_visibility(styles_xml, namespaces) -> _WordStyleVisibility:
    if not styles_xml:
        return _WordStyleVisibility(
            frozenset(),
            frozenset(),
            frozenset(),
            frozenset(),
            False,
            False,
            "",
        )
    try:
        styles_root = ElementTree.fromstring(styles_xml)
    except ElementTree.ParseError as exc:
        raise RuntimeError(f"Word 样式结构损坏：{exc}") from exc

    hidden_by_style = {}
    color_state_by_style = {}
    invalid_conditional_by_style = {}
    based_on = {}
    default_style_ids = set()
    default_table_style_id = ""
    for style in styles_root.findall("./w:style", namespaces):
        style_id = style.get(f"{{{_WORD_NAMESPACE}}}styleId", "")
        if not style_id:
            continue
        style_type = str(
            style.get(f"{{{_WORD_NAMESPACE}}}type", "")
        ).strip().lower()
        run_properties = (
            tuple(style.findall("./w:rPr", namespaces))
            + tuple(style.findall("./w:pPr/w:rPr", namespaces))
        )
        hidden_by_style[style_id] = any(
            _word_run_properties_hide_text(properties, namespaces)
            or _word_properties_make_text_invisible(properties, namespaces)
            for properties in run_properties
        )
        direct_color_state = next(
            (
                state
                for state in (
                    _word_properties_color_state(properties, namespaces)
                    for properties in reversed(run_properties)
                )
                if state is not None
            ),
            None,
        )
        style_has_dark_background = any(
            _word_properties_have_dark_background(properties, namespaces)
            for properties in (
                tuple(style.findall("./w:tcPr", namespaces))
                + tuple(style.findall("./w:pPr", namespaces))
                + tuple(style.findall("./w:tblPr", namespaces))
            )
        )
        if (
            direct_color_state is True
            and style_type == "table"
            and style_has_dark_background
        ):
            direct_color_state = False
        color_state_by_style[style_id] = direct_color_state

        invalid_conditional = False
        for conditional in style.findall("./w:tblStylePr", namespaces):
            conditional_run_properties = (
                tuple(conditional.findall("./w:rPr", namespaces))
                + tuple(conditional.findall("./w:pPr/w:rPr", namespaces))
            )
            if any(
                _word_run_properties_hide_text(properties, namespaces)
                or _word_properties_make_text_invisible(properties, namespaces)
                for properties in conditional_run_properties
            ):
                invalid_conditional = True
                break
            conditional_near_white = any(
                _word_properties_use_near_white(properties, namespaces)
                for properties in conditional_run_properties
            )
            conditional_dark_background = any(
                _word_properties_have_dark_background(properties, namespaces)
                for properties in (
                    tuple(conditional.findall("./w:tcPr", namespaces))
                    + tuple(conditional.findall("./w:pPr", namespaces))
                    + tuple(conditional.findall("./w:tblPr", namespaces))
                )
            )
            if conditional_near_white and not conditional_dark_background:
                invalid_conditional = True
                break
        invalid_conditional_by_style[style_id] = invalid_conditional

        parent = style.find("./w:basedOn", namespaces)
        based_on[style_id] = (
            parent.get(f"{{{_WORD_NAMESPACE}}}val", "")
            if parent is not None
            else ""
        )
        default_value = str(
            style.get(f"{{{_WORD_NAMESPACE}}}default", "0")
        ).strip().lower()
        if default_value not in {"0", "false", "off", "none"}:
            default_style_ids.add(style_id)
            if style_type == "table":
                default_table_style_id = style_id

    def resolve_flag(style_id, direct_values, resolved, active=frozenset()):
        if not style_id:
            return False
        if style_id in resolved:
            return resolved[style_id]
        if style_id in active:
            return False
        result = bool(direct_values.get(style_id)) or resolve_flag(
            based_on.get(style_id, ""),
            direct_values,
            resolved,
            active | {style_id},
        )
        resolved[style_id] = result
        return result

    def resolve_value(style_id, direct_values, resolved, active=frozenset()):
        if not style_id:
            return None
        if style_id in resolved:
            return resolved[style_id]
        if style_id in active:
            return None
        direct_value = direct_values.get(style_id)
        result = (
            direct_value
            if direct_value is not None
            else resolve_value(
                based_on.get(style_id, ""),
                direct_values,
                resolved,
                active | {style_id},
            )
        )
        resolved[style_id] = result
        return result

    resolved_hidden = {}
    resolved_color_state = {}
    resolved_invalid_conditional = {}
    hidden_style_ids = frozenset(
        style_id
        for style_id in hidden_by_style
        if resolve_flag(style_id, hidden_by_style, resolved_hidden)
    )
    explicit_color_style_ids = frozenset(
        style_id
        for style_id in color_state_by_style
        if resolve_value(
            style_id,
            color_state_by_style,
            resolved_color_state,
        ) is not None
    )
    near_white_style_ids = frozenset(
        style_id
        for style_id in explicit_color_style_ids
        if resolved_color_state.get(style_id) is True
    )
    invalid_conditional_table_style_ids = frozenset(
        style_id
        for style_id in invalid_conditional_by_style
        if resolve_flag(
            style_id,
            invalid_conditional_by_style,
            resolved_invalid_conditional,
        )
    )
    default_properties = tuple(
        styles_root.findall("./w:docDefaults/w:rPrDefault/w:rPr", namespaces)
    )
    default_hidden = any(
        _word_run_properties_hide_text(properties, namespaces)
        or _word_properties_make_text_invisible(properties, namespaces)
        for properties in default_properties
    ) or any(style_id in hidden_style_ids for style_id in default_style_ids)
    default_near_white = any(
        _word_properties_use_near_white(properties, namespaces)
        for properties in default_properties
    ) or any(style_id in near_white_style_ids for style_id in default_style_ids)
    return _WordStyleVisibility(
        hidden_style_ids,
        near_white_style_ids,
        explicit_color_style_ids,
        invalid_conditional_table_style_ids,
        default_hidden,
        default_near_white,
        default_table_style_id,
    )

def _hidden_word_style_ids(styles_xml, namespaces) -> frozenset[str]:
    return _word_style_visibility(styles_xml, namespaces).hidden_style_ids

def _word_hex_color(value):
    normalized = str(value or "").strip().lstrip("#")
    if len(normalized) != 6:
        return None
    try:
        return tuple(int(normalized[index : index + 2], 16) for index in (0, 2, 4))
    except ValueError:
        return None


def _word_properties_color_state(properties, namespaces):
    if properties is None:
        return None
    color = properties.find("./w:color", namespaces)
    if color is None:
        return None
    value = str(color.get(f"{{{_WORD_NAMESPACE}}}val", "")).strip().lower()
    rgb = _word_hex_color(value)
    if rgb is not None:
        return min(rgb) >= 245
    theme_color = str(
        color.get(f"{{{_WORD_NAMESPACE}}}themeColor", "")
    ).strip().lower()
    if theme_color in {"background1", "light1", "background2", "light2"}:
        return True
    if theme_color in {"text1", "dark1", "text2", "dark2"}:
        return False
    # An explicit color that cannot be resolved is not safe evidence of visible text.
    return True


def _word_properties_background_rgb(properties, namespaces):
    if properties is None:
        return None
    shading = properties.find("./w:shd", namespaces)
    if shading is not None:
        rgb = _word_hex_color(
            shading.get(f"{{{_WORD_NAMESPACE}}}fill", "")
        )
        if rgb is not None:
            return rgb
    highlight = properties.find("./w:highlight", namespaces)
    if highlight is None:
        return None
    highlight_name = str(
        highlight.get(f"{{{_WORD_NAMESPACE}}}val", "")
    ).strip().lower()
    return {
        "black": (0, 0, 0),
        "blue": (0, 0, 255),
        "cyan": (0, 255, 255),
        "darkblue": (0, 0, 139),
        "darkcyan": (0, 139, 139),
        "darkgray": (128, 128, 128),
        "darkgreen": (0, 100, 0),
        "darkmagenta": (139, 0, 139),
        "darkred": (139, 0, 0),
        "darkyellow": (128, 128, 0),
        "green": (0, 128, 0),
        "lightgray": (211, 211, 211),
        "magenta": (255, 0, 255),
        "red": (255, 0, 0),
        "white": (255, 255, 255),
        "yellow": (255, 255, 0),
    }.get(highlight_name)


def _relative_luminance(rgb) -> float:
    channels = []
    for value in rgb:
        normalized = max(0.0, min(1.0, float(value) / 255.0))
        channels.append(
            normalized / 12.92
            if normalized <= 0.04045
            else ((normalized + 0.055) / 1.055) ** 2.4
        )
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def _word_properties_have_dark_background(properties, namespaces) -> bool:
    background = _word_properties_background_rgb(properties, namespaces)
    if background is None:
        return False
    foreground_luminance = _relative_luminance((245, 245, 245))
    background_luminance = _relative_luminance(background)
    lighter = max(foreground_luminance, background_luminance)
    darker = min(foreground_luminance, background_luminance)
    return (lighter + 0.05) / (darker + 0.05) >= 3.0

def _word_properties_make_text_invisible(properties, namespaces) -> bool:
    if properties is None:
        return False
    for size_name in ("sz", "szCs"):
        size = properties.find(f"./w:{size_name}", namespaces)
        if size is None:
            continue
        try:
            if int(size.get(f"{{{_WORD_NAMESPACE}}}val", "0")) <= 2:
                return True
        except (TypeError, ValueError):
            return True
    text_scale = properties.find("./w:w", namespaces)
    if text_scale is not None:
        try:
            if int(text_scale.get(f"{{{_WORD_NAMESPACE}}}val", "100")) <= 1:
                return True
        except (TypeError, ValueError):
            return True
    for element in properties.iter():
        local_name = element.tag.rsplit("}", 1)[-1]
        if local_name == "noFill":
            return True
        if local_name == "alpha":
            try:
                if int(element.get("val", "100000")) <= 1000:
                    return True
            except (TypeError, ValueError):
                return True
    return False


def _word_properties_use_near_white(properties, namespaces) -> bool:
    return _word_properties_color_state(properties, namespaces) is True

def _word_text_has_dark_background(text_node, parent_map, namespaces) -> bool:
    current = parent_map.get(text_node)
    while current is not None:
        properties = None
        if current.tag == f"{{{_WORD_NAMESPACE}}}r":
            properties = current.find("./w:rPr", namespaces)
        elif current.tag == f"{{{_WORD_NAMESPACE}}}p":
            properties = current.find("./w:pPr", namespaces)
        elif current.tag == f"{{{_WORD_NAMESPACE}}}tc":
            properties = current.find("./w:tcPr", namespaces)
        if _word_properties_have_dark_background(properties, namespaces):
            return True
        current = parent_map.get(current)
    return False

def _word_text_node_is_hidden(
    text_node,
    parent_map,
    style_visibility,
    namespaces,
) -> bool:
    current = parent_map.get(text_node)
    if style_visibility.default_hidden:
        return True
    near_white_text = None
    while current is not None:
        if current.tag == f"{{{_WORD_NAMESPACE}}}r":
            properties = current.find("./w:rPr", namespaces)
            if (
                _word_run_properties_hide_text(properties, namespaces)
                or _word_properties_make_text_invisible(properties, namespaces)
            ):
                return True
            direct_color_state = _word_properties_color_state(
                properties,
                namespaces,
            )
            if near_white_text is None and direct_color_state is not None:
                near_white_text = direct_color_state
            style = (
                properties.find("./w:rStyle", namespaces)
                if properties is not None
                else None
            )
            style_id = (
                style.get(f"{{{_WORD_NAMESPACE}}}val", "")
                if style is not None
                else ""
            )
            if style_id in style_visibility.hidden_style_ids:
                return True
            if (
                near_white_text is None
                and style_id in style_visibility.explicit_color_style_ids
            ):
                near_white_text = (
                    style_id in style_visibility.near_white_style_ids
                )
        elif current.tag == f"{{{_WORD_NAMESPACE}}}p":
            properties = current.find("./w:pPr", namespaces)
            run_properties = (
                properties.find("./w:rPr", namespaces)
                if properties is not None
                else None
            )
            if (
                _word_run_properties_hide_text(run_properties, namespaces)
                or _word_properties_make_text_invisible(
                    run_properties,
                    namespaces,
                )
            ):
                return True
            paragraph_color_state = _word_properties_color_state(
                run_properties,
                namespaces,
            )
            if near_white_text is None and paragraph_color_state is not None:
                near_white_text = paragraph_color_state
            style = (
                properties.find("./w:pStyle", namespaces)
                if properties is not None
                else None
            )
            style_id = (
                style.get(f"{{{_WORD_NAMESPACE}}}val", "")
                if style is not None
                else ""
            )
            if style_id in style_visibility.hidden_style_ids:
                return True
            if (
                near_white_text is None
                and style_id in style_visibility.explicit_color_style_ids
            ):
                near_white_text = (
                    style_id in style_visibility.near_white_style_ids
                )
        elif current.tag == f"{{{_WORD_NAMESPACE}}}tbl":
            properties = current.find("./w:tblPr", namespaces)
            style = (
                properties.find("./w:tblStyle", namespaces)
                if properties is not None
                else None
            )
            style_id = (
                style.get(f"{{{_WORD_NAMESPACE}}}val", "")
                if style is not None
                else ""
            )
            if style_id in style_visibility.hidden_style_ids:
                return True
            if (
                near_white_text is None
                and style_id in style_visibility.explicit_color_style_ids
            ):
                near_white_text = (
                    style_id in style_visibility.near_white_style_ids
                )
        current = parent_map.get(current)
    if near_white_text is None:
        near_white_text = style_visibility.default_near_white
    return bool(
        near_white_text
        and not _word_text_has_dark_background(
            text_node,
            parent_map,
            namespaces,
        )
    )


def _word_text_node_is_discarded(text_node, parent_map) -> bool:
    discarded_tags = {
        f"{{{_WORD_NAMESPACE}}}del",
        f"{{{_WORD_NAMESPACE}}}moveFrom",
    }
    current = parent_map.get(text_node)
    while current is not None:
        if current.tag in discarded_tags:
            return True
        current = parent_map.get(current)
    return False


def _nearest_word_ancestor(element, parent_map, local_name):
    expected_tag = f"{{{_WORD_NAMESPACE}}}{local_name}"
    current = element
    while current is not None:
        if current.tag == expected_tag:
            return current
        current = parent_map.get(current)
    return None


def _word_text_node_belongs_to_container(text_node, container, parent_map) -> bool:
    table_tag = f"{{{_WORD_NAMESPACE}}}tbl"
    row_tag = f"{{{_WORD_NAMESPACE}}}tr"
    cell_tag = f"{{{_WORD_NAMESPACE}}}tc"
    paragraph_tag = f"{{{_WORD_NAMESPACE}}}p"
    run_tag = f"{{{_WORD_NAMESPACE}}}r"
    forbidden_tags = {
        f"{{{_WORD_NAMESPACE}}}drawing",
        f"{{{_WORD_NAMESPACE}}}pict",
        f"{{{_WORD_NAMESPACE}}}txbxContent",
        f"{{{_WORD_NAMESPACE}}}object",
        f"{{{_WORD_NAMESPACE}}}del",
        f"{{{_WORD_NAMESPACE}}}moveFrom",
    }
    owner_table = (
        container
        if container.tag == table_tag
        else _nearest_word_ancestor(container, parent_map, "tbl")
    )
    owner_cell = container if container.tag == cell_tag else None
    nearest_table = None
    nearest_row = None
    nearest_cell = None
    nearest_paragraph = None
    nearest_run = None
    current = parent_map.get(text_node)
    while current is not None:
        if current.tag in forbidden_tags:
            return False
        if nearest_run is None and current.tag == run_tag:
            nearest_run = current
        if nearest_paragraph is None and current.tag == paragraph_tag:
            nearest_paragraph = current
        if nearest_cell is None and current.tag == cell_tag:
            nearest_cell = current
        if nearest_row is None and current.tag == row_tag:
            nearest_row = current
        if nearest_table is None and current.tag == table_tag:
            nearest_table = current
        current = parent_map.get(current)
    if any(
        element is None
        for element in (
            nearest_run,
            nearest_paragraph,
            nearest_cell,
            nearest_row,
            nearest_table,
        )
    ):
        return False
    if owner_table is None or nearest_table is not owner_table:
        return False
    if owner_cell is not None and nearest_cell is not owner_cell:
        return False
    return True

def _visible_word_text_nodes(
    container,
    parent_map,
    style_visibility,
    namespaces,
):
    return tuple(
        text_node
        for text_node in container.findall(".//w:t", namespaces)
        if _word_text_node_belongs_to_container(
            text_node,
            container,
            parent_map,
        )
        and not _word_text_node_is_hidden(
            text_node,
            parent_map,
            style_visibility,
            namespaces,
        )
    )

def _word_attribute(element, name, default=None):
    if element is None:
        return default
    value = element.get(f"{{{_WORD_NAMESPACE}}}{name}")
    if value is None:
        value = element.get(name)
    return default if value is None else value


def _word_integer_attribute(element, name):
    value = _word_attribute(element, name)
    if value is None:
        return None
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return None


def _word_width_twips(
    element,
    *,
    available_width_twips,
    default_type="dxa",
    allow_auto=False,
):
    if element is None:
        return None
    raw_width = _word_attribute(element, "w")
    if raw_width is None:
        return None
    width_type = str(
        _word_attribute(element, "type", default_type)
    ).strip().lower()
    if width_type == "auto":
        return None if allow_auto else -1.0
    if width_type == "nil":
        return -1.0
    try:
        width = int(str(raw_width).strip())
    except (TypeError, ValueError):
        return -1.0
    if width < 0:
        return -1.0
    if width_type == "pct":
        return float(available_width_twips) * width / 5000.0
    if width_type not in {"dxa", ""}:
        return -1.0
    return float(width)


def _word_width_is_negligible(
    element,
    *,
    available_width_twips,
    minimum_twips,
    default_type="dxa",
    allow_auto=False,
    require_value=False,
) -> bool:
    width = _word_width_twips(
        element,
        available_width_twips=available_width_twips,
        default_type=default_type,
        allow_auto=allow_auto,
    )
    if width is None:
        return bool(require_value and not allow_auto)
    return width <= float(minimum_twips)

def _word_page_dimensions(root, namespaces):
    dimensions = []
    for page_size in root.findall(".//w:sectPr/w:pgSz", namespaces):
        width = _word_integer_attribute(page_size, "w")
        height = _word_integer_attribute(page_size, "h")
        if width and height and width > 0 and height > 0:
            dimensions.append((width, height))
    return tuple(dimensions) or ((12240, 15840),)


def _word_minimum_usable_page_width(root, namespaces) -> int:
    usable_widths = []
    for section in root.findall(".//w:sectPr", namespaces):
        page_size = section.find("./w:pgSz", namespaces)
        page_width = _word_integer_attribute(page_size, "w") or 12240
        margins = section.find("./w:pgMar", namespaces)
        left_margin = _word_integer_attribute(margins, "left")
        right_margin = _word_integer_attribute(margins, "right")
        left_margin = 1440 if left_margin is None else max(0, left_margin)
        right_margin = 1440 if right_margin is None else max(0, right_margin)
        usable_width = page_width - left_margin - right_margin
        if usable_width > 0:
            usable_widths.append(usable_width)
    return min(usable_widths) if usable_widths else 9360

def _word_text_size_half_points(text_node, parent_map, namespaces) -> int:
    current = parent_map.get(text_node)
    while current is not None:
        properties = None
        if current.tag == f"{{{_WORD_NAMESPACE}}}r":
            properties = current.find("./w:rPr", namespaces)
        elif current.tag == f"{{{_WORD_NAMESPACE}}}p":
            paragraph_properties = current.find("./w:pPr", namespaces)
            properties = (
                paragraph_properties.find("./w:rPr", namespaces)
                if paragraph_properties is not None
                else None
            )
        if properties is not None:
            sizes = []
            for size_name in ("sz", "szCs"):
                size = properties.find(f"./w:{size_name}", namespaces)
                value = _word_integer_attribute(size, "val")
                if value is not None and value > 0:
                    sizes.append(value)
            if sizes:
                return max(sizes)
        current = parent_map.get(current)
    return 22


def _word_table_has_invalid_layout(
    table,
    root,
    parent_map,
    style_visibility,
    namespaces,
) -> bool:
    forbidden_tags = {
        f"{{{_WORD_NAMESPACE}}}drawing",
        f"{{{_WORD_NAMESPACE}}}pict",
        f"{{{_WORD_NAMESPACE}}}txbxContent",
        f"{{{_WORD_NAMESPACE}}}object",
    }
    if any(element.tag in forbidden_tags for element in table.iter()):
        return True
    current = parent_map.get(table)
    while current is not None:
        if current.tag in forbidden_tags:
            return True
        current = parent_map.get(current)

    for text_node in table.findall(".//w:t", namespaces):
        if _nearest_word_ancestor(text_node, parent_map, "tbl") is not table:
            continue
        if not _word_text_node_belongs_to_container(
            text_node,
            table,
            parent_map,
        ):
            return True
        if _word_text_node_is_hidden(
            text_node,
            parent_map,
            style_visibility,
            namespaces,
        ):
            return True

    table_style = table.find("./w:tblPr/w:tblStyle", namespaces)
    table_style_id = str(
        _word_attribute(
            table_style,
            "val",
            style_visibility.default_table_style_id,
        )
    )
    if (
        table_style_id
        in style_visibility.invalid_conditional_table_style_ids
    ):
        return True

    # Floating tables require section-, anchor- and wrapping-aware layout math.
    # Conversion output cannot be certified visible without that full model.
    if table.find("./w:tblPr/w:tblpPr", namespaces) is not None:
        return True

    available_page_width = _word_minimum_usable_page_width(root, namespaces)
    grid_columns = tuple(
        table.findall("./w:tblGrid/w:gridCol", namespaces)
    )
    if not grid_columns:
        return True
    grid_widths = []
    for column in grid_columns:
        width = _word_width_twips(
            column,
            available_width_twips=available_page_width,
        )
        if width is None or width <= 20:
            return True
        grid_widths.append(width)
    grid_total_width = sum(grid_widths)
    if grid_total_width < max(120, 20 * len(grid_columns)):
        return True

    table_width = table.find("./w:tblPr/w:tblW", namespaces)
    if table_width is not None:
        if _word_attribute(table_width, "w") is None:
            return True
        if _word_width_is_negligible(
            table_width,
            available_width_twips=available_page_width,
            minimum_twips=120,
            default_type="auto",
            allow_auto=True,
        ):
            return True
    for cell_width in table.findall(
        "./w:tr/w:tc/w:tcPr/w:tcW",
        namespaces,
    ):
        if _word_attribute(cell_width, "w") is None:
            return True
        if _word_width_is_negligible(
            cell_width,
            available_width_twips=grid_total_width,
            minimum_twips=20,
            default_type="auto",
            allow_auto=True,
        ):
            return True

    for row in table.findall("./w:tr", namespaces):
        grid_before = row.find("./w:trPr/w:gridBefore", namespaces)
        grid_after = row.find("./w:trPr/w:gridAfter", namespaces)
        leading_columns = _word_integer_attribute(grid_before, "val")
        trailing_columns = _word_integer_attribute(grid_after, "val")
        if grid_before is not None and leading_columns is None:
            return True
        if grid_after is not None and trailing_columns is None:
            return True
        leading_columns = 0 if leading_columns is None else leading_columns
        trailing_columns = 0 if trailing_columns is None else trailing_columns
        if leading_columns < 0 or trailing_columns < 0:
            return True
        logical_column_count = leading_columns + trailing_columns
        cells = tuple(row.findall("./w:tc", namespaces))
        if not cells:
            return True
        for cell in cells:
            grid_span = cell.find("./w:tcPr/w:gridSpan", namespaces)
            span = _word_integer_attribute(grid_span, "val")
            if grid_span is not None and span is None:
                return True
            span = 1 if span is None else span
            if span < 1:
                return True
            logical_column_count += span
        if logical_column_count != len(grid_columns):
            return True

        row_heights = tuple(
            row.findall("./w:trPr/w:trHeight", namespaces)
        )
        if len(row_heights) > 1:
            return True
        if not row_heights:
            continue
        row_height = row_heights[0]
        height_rule = str(
            _word_attribute(row_height, "hRule", "atLeast")
        ).strip().lower()
        if height_rule != "exact":
            continue
        height = _word_integer_attribute(row_height, "val")
        if height is None or height <= 40:
            return True
        visible_nodes = _visible_word_text_nodes(
            row,
            parent_map,
            style_visibility,
            namespaces,
        )
        if not visible_nodes:
            continue
        largest_font_size = max(
            _word_text_size_half_points(
                text_node,
                parent_map,
                namespaces,
            )
            for text_node in visible_nodes
        )
        minimum_visible_height = max(40, largest_font_size * 7)
        if height < minimum_visible_height:
            return True
    return False

def _large_page_drawing_count(root, namespaces) -> int:
    page_sizes = []
    for page_size in root.findall(".//w:sectPr/w:pgSz", namespaces):
        try:
            width = int(page_size.get(f"{{{_WORD_NAMESPACE}}}w", "0"))
            height = int(page_size.get(f"{{{_WORD_NAMESPACE}}}h", "0"))
        except (TypeError, ValueError):
            continue
        if width > 0 and height > 0:
            page_sizes.append((width, height))
    if not page_sizes:
        page_sizes.append((12240, 15840))

    large_count = 0
    drawings = root.findall(".//wp:inline", namespaces) + root.findall(
        ".//wp:anchor",
        namespaces,
    )
    for drawing in drawings:
        extent = drawing.find("./wp:extent", namespaces)
        if extent is None:
            continue
        try:
            width_twips = int(extent.get("cx", "0")) / 635
            height_twips = int(extent.get("cy", "0")) / 635
        except (TypeError, ValueError):
            continue
        if width_twips <= 0 or height_twips <= 0:
            continue
        for page_width, page_height in page_sizes:
            width_ratio = width_twips / page_width
            height_ratio = height_twips / page_height
            area_ratio = width_ratio * height_ratio
            if (
                width_ratio >= 0.72
                and height_ratio >= 0.45
                and area_ratio >= 0.36
            ):
                large_count += 1
                break
    return large_count


def _word_cell_has_explicit_border(cell, names, namespaces) -> bool:
    borders = cell.find("./w:tcPr/w:tcBorders", namespaces)
    if borders is None:
        return False
    for name in names:
        edge = borders.find(f"./w:{name}", namespaces)
        if edge is None:
            continue
        value = str(_word_attribute(edge, "val", "")).strip().casefold()
        if value and value not in {"nil", "none"}:
            return True
    return False


def _word_logical_cell_border_edges(physical_cells, namespaces):
    """Aggregate a logical cell's explicit outer borders across vMerge pieces."""

    physical_cells = tuple(physical_cells)
    if not physical_cells:
        return (False, False, False, False)
    return (
        _word_cell_has_explicit_border(physical_cells[0], ("top",), namespaces),
        all(
            _word_cell_has_explicit_border(cell, ("left", "start"), namespaces)
            for cell in physical_cells
        ),
        _word_cell_has_explicit_border(physical_cells[-1], ("bottom",), namespaces),
        all(
            _word_cell_has_explicit_border(cell, ("right", "end"), namespaces)
            for cell in physical_cells
        ),
    )

def inspect_editable_docx_tables(path: str | PathLike[str]) -> EditableDocxTableSummary:
    """Inspect WordprocessingML instead of trusting a visually similar page image."""

    docx_path = Path(path)
    try:
        from docx import Document

        Document(str(docx_path))
    except Exception as exc:
        raise RuntimeError(f"无法打开 Word 文档“{docx_path}”：{exc}") from exc

    try:
        with ZipFile(docx_path) as archive:
            try:
                settings_xml = archive.read("word/settings.xml")
            except KeyError:
                settings_xml = b""
            try:
                styles_xml = archive.read("word/styles.xml")
            except KeyError:
                styles_xml = b""
            document_xml = archive.read("word/document.xml")
    except (BadZipFile, KeyError, OSError) as exc:
        raise RuntimeError(f"无法检查 Word 表格结构“{docx_path}”：{exc}") from exc

    try:
        root = ElementTree.fromstring(document_xml)
    except ElementTree.ParseError as exc:
        raise RuntimeError(f"Word 文档结构损坏“{docx_path}”：{exc}") from exc

    namespaces = {"w": _WORD_NAMESPACE, "wp": _DRAWING_NAMESPACE}
    parent_map = {
        child: parent
        for parent in root.iter()
        for child in parent
    }
    style_visibility = _word_style_visibility(styles_xml, namespaces)
    visible_document_text = "".join(
        text_node.text or ""
        for text_node in root.findall(".//w:t", namespaces)
        if not _word_text_node_is_discarded(text_node, parent_map)
        and not _word_text_node_is_hidden(
            text_node,
            parent_map,
            style_visibility,
            namespaces,
        )
    )
    visible_document_text_character_count = sum(
        not character.isspace() for character in visible_document_text
    )
    document_protected = False
    if settings_xml:
        try:
            settings_root = ElementTree.fromstring(settings_xml)
        except ElementTree.ParseError as exc:
            raise RuntimeError(f"Word 设置结构损坏“{docx_path}”：{exc}") from exc
        protection = settings_root.find(".//w:documentProtection", namespaces)
        if protection is not None:
            enforcement = protection.get(f"{{{_WORD_NAMESPACE}}}enforcement", "1")
            document_protected = enforcement not in {"0", "false", "False", "off"}
        if settings_root.find(".//w:writeProtection", namespaces) is not None:
            document_protected = True

    tables = root.findall(".//w:tbl", namespaces)
    locked_content_controls = []
    locked_values = {"contentLocked", "sdtContentLocked"}
    for content_control in root.findall(".//w:sdt", namespaces):
        lock = content_control.find("./w:sdtPr/w:lock", namespaces)
        if lock is None:
            continue
        lock_value = lock.get(f"{{{_WORD_NAMESPACE}}}val", "")
        if lock_value in locked_values:
            locked_content_controls.append(content_control)

    locked_table_ids = set()
    for table in tables:
        table_descendant_ids = {id(element) for element in table.iter()}
        if any(
            id(content_control) in table_descendant_ids
            or any(element is table for element in content_control.iter())
            for content_control in locked_content_controls
        ):
            locked_table_ids.add(id(table))

    invalid_layout_table_ids = {
        id(table)
        for table in tables
        if _word_table_has_invalid_layout(
            table,
            root,
            parent_map,
            style_visibility,
            namespaces,
        )
    }

    table_text_character_count = 0
    table_texts = []
    table_shapes = []
    table_cell_counts = []
    table_raw_cell_counts = []
    table_cell_matrices = []
    table_cell_spans = []
    table_grid_widths = []
    table_cell_border_edges = []
    for table in tables:
        rows = table.findall("./w:tr", namespaces)
        grid_widths = tuple(
            int(_word_integer_attribute(column, "w") or 0)
            for column in table.findall("./w:tblGrid/w:gridCol", namespaces)
        )
        declared_grid_column_count = len(grid_widths)
        table_grid_widths.append(grid_widths)

        logical_cell_count = 0
        raw_cell_count = 0
        matrix_rows = []
        span_records: list[list[int]] = []
        span_cells: list[list[ElementTree.Element]] = []
        active_vertical_merges: dict[tuple[int, int], int] = {}
        topology_invalid = False

        for row_index, row in enumerate(rows):
            grid_before = row.find("./w:trPr/w:gridBefore", namespaces)
            leading_columns = _word_integer_attribute(grid_before, "val")
            leading_columns = 0 if leading_columns is None else max(0, leading_columns)
            row_values = [None] * leading_columns
            grid_column = leading_columns
            current_vertical_merges: dict[tuple[int, int], int] = {}
            cells = row.findall("./w:tc", namespaces)
            raw_cell_count += len(cells)

            for cell in cells:
                grid_span = cell.find("./w:tcPr/w:gridSpan", namespaces)
                span = _word_integer_attribute(grid_span, "val")
                span = 1 if span is None else max(1, span)
                start_column = grid_column
                merge_key = (start_column, span)

                vertical_merge = cell.find("./w:tcPr/w:vMerge", namespaces)
                merge_value = ""
                if vertical_merge is not None:
                    merge_value = str(
                        _word_attribute(vertical_merge, "val", "continue")
                    ).strip().lower()
                    if merge_value not in {"restart", "continue"}:
                        topology_invalid = True

                visible_text = "".join(
                    text_node.text or ""
                    for text_node in _visible_word_text_nodes(
                        cell,
                        parent_map,
                        style_visibility,
                        namespaces,
                    )
                )
                is_merge_continuation = (
                    vertical_merge is not None and merge_value != "restart"
                )
                if is_merge_continuation:
                    span_index = active_vertical_merges.get(merge_key)
                    if span_index is None:
                        topology_invalid = True
                    else:
                        span_records[span_index][2] += 1
                        span_cells[span_index].append(cell)
                        current_vertical_merges[merge_key] = span_index
                    if visible_text.strip():
                        topology_invalid = True
                    cell_text = None
                else:
                    span_index = len(span_records)
                    span_records.append([row_index, start_column, 1, span])
                    span_cells.append([cell])
                    logical_cell_count += 1
                    cell_text = visible_text
                    if vertical_merge is not None:
                        current_vertical_merges[merge_key] = span_index

                row_values.append(cell_text)
                row_values.extend([None] * (span - 1))
                grid_column += span

            grid_after = row.find("./w:trPr/w:gridAfter", namespaces)
            trailing_columns = _word_integer_attribute(grid_after, "val")
            trailing_columns = 0 if trailing_columns is None else max(0, trailing_columns)
            if trailing_columns:
                row_values.extend([None] * trailing_columns)
                grid_column += trailing_columns

            if declared_grid_column_count and grid_column != declared_grid_column_count:
                topology_invalid = True
            matrix_width = max(declared_grid_column_count, grid_column)
            matrix_rows.append(
                tuple(row_values + [None] * (matrix_width - len(row_values)))
            )
            active_vertical_merges = current_vertical_merges

        matrix_width = max(
            declared_grid_column_count,
            max((len(row) for row in matrix_rows), default=0),
        )
        table_matrix = tuple(
            tuple(row) + (None,) * (matrix_width - len(row))
            for row in matrix_rows
        )
        table_text = "".join(
            str(cell_text or "")
            for row_values in table_matrix
            for cell_text in row_values
            if cell_text is not None
        )
        span_bundles = tuple(
            sorted(
                (
                    (tuple(record), tuple(physical_cells))
                    for record, physical_cells in zip(span_records, span_cells)
                ),
                key=lambda item: item[0],
            )
        )
        spans = tuple(record for record, _physical_cells in span_bundles)
        border_edges = tuple(
            _word_logical_cell_border_edges(physical_cells, namespaces)
            for _record, physical_cells in span_bundles
        )
        if (
            logical_cell_count != len(spans)
            or len(border_edges) != len(spans)
        ):
            topology_invalid = True
        if topology_invalid:
            invalid_layout_table_ids.add(id(table))

        table_shapes.append((len(rows), matrix_width))
        table_cell_counts.append(logical_cell_count)
        table_raw_cell_counts.append(raw_cell_count)
        table_cell_matrices.append(table_matrix)
        table_cell_spans.append(spans)
        table_cell_border_edges.append(border_edges)
        table_texts.append(table_text)
        table_text_character_count += sum(
            not character.isspace() for character in table_text
        )
    return EditableDocxTableSummary(
        table_count=len(tables),
        row_count=len(root.findall(".//w:tbl/w:tr", namespaces)),
        cell_count=len(root.findall(".//w:tbl/w:tr/w:tc", namespaces)),
        document_protected=document_protected,
        locked_content_control_table_count=len(locked_table_ids),
        table_text_character_count=table_text_character_count,
        table_texts=tuple(table_texts),
        table_cell_matrices=tuple(table_cell_matrices),
        table_shapes=tuple(table_shapes),
        table_cell_counts=tuple(table_cell_counts),
        table_cell_spans=tuple(table_cell_spans),
        table_grid_widths=tuple(table_grid_widths),
        table_raw_cell_counts=tuple(table_raw_cell_counts),
        table_cell_border_edges=tuple(table_cell_border_edges),
        drawing_count=len(root.findall(".//wp:inline", namespaces))
        + len(root.findall(".//wp:anchor", namespaces)),
        large_page_drawing_count=_large_page_drawing_count(root, namespaces),
        invalid_layout_table_count=len(invalid_layout_table_ids),
        visible_text_character_count=visible_document_text_character_count,
        visible_text=visible_document_text,
    )

_PDF_XREF_REFERENCE = re.compile(r"(?<!\d)(\d+)\s+\d+\s+R")
_PDF_NAME_PAIR = re.compile(
    r"/([^\s/<>\[\]()%]+)\s*/([^\s/<>\[\]()%]+)"
)
_PDF_STRUCTURE_ROLE = re.compile(
    r"/S\s*/([^\s/<>\[\]()%]+)"
)
_PDF_NAME_ESCAPE = re.compile(r"#([0-9A-Fa-f]{2})")
_PDF_TABLE_ROLES = frozenset(
    {"Table", "TR", "TH", "TD", "THead", "TBody", "TFoot"}
)


def _decode_pdf_name(value) -> str:
    name = str(value or "").lstrip("/")
    return _PDF_NAME_ESCAPE.sub(
        lambda match: chr(int(match.group(1), 16)),
        name,
    )


def _table_roles_from_role_map(role_map_value) -> set[str]:
    mappings = {
        _decode_pdf_name(alias): _decode_pdf_name(target)
        for alias, target in _PDF_NAME_PAIR.findall(str(role_map_value or ""))
    }
    table_roles = set(_PDF_TABLE_ROLES)
    for _iteration in range(len(mappings) + 1):
        aliases = {
            alias
            for alias, target in mappings.items()
            if target in table_roles
        }
        if aliases <= table_roles:
            break
        table_roles.update(aliases)
    return table_roles


def _value_has_table_structure_role(value, table_roles) -> bool:
    return any(
        _decode_pdf_name(role) in table_roles
        for role in _PDF_STRUCTURE_ROLE.findall(str(value or ""))
    )


def _document_has_tagged_table_structure(document, node_limit=20000) -> bool:
    """Treat PDF table structure roles as high-confidence table signals."""

    pdf_catalog = getattr(document, "pdf_catalog", None)
    xref_get_key = getattr(document, "xref_get_key", None)
    xref_object = getattr(document, "xref_object", None)
    if not (
        callable(pdf_catalog)
        and callable(xref_get_key)
        and callable(xref_object)
    ):
        return False
    catalog = pdf_catalog()
    if not isinstance(catalog, int) or catalog <= 0:
        return False

    root_kind, root_value = xref_get_key(catalog, "StructTreeRoot")
    if root_kind in {"null", "none"}:
        return False
    root_references = [
        int(value)
        for value in _PDF_XREF_REFERENCE.findall(str(root_value))
    ]

    role_map_value = ""
    if root_kind == "xref" and root_references:
        role_kind, role_value = xref_get_key(
            root_references[0],
            "RoleMap",
        )
        if role_kind == "xref":
            role_references = _PDF_XREF_REFERENCE.findall(str(role_value))
            if role_references:
                role_map_value = xref_object(
                    int(role_references[0]),
                    compressed=False,
                )
        elif role_kind not in {"null", "none"}:
            role_map_value = role_value
    elif root_kind in {"dict", "array"}:
        role_map_value = root_value

    table_roles = _table_roles_from_role_map(role_map_value)
    if _value_has_table_structure_role(root_value, table_roles):
        return True
    if not root_references:
        raise RuntimeError("PDF 标签结构存在，但 StructTreeRoot 无法解析")

    queue = list(root_references)
    visited = set()
    while queue:
        xref = queue.pop(0)
        if xref in visited:
            continue
        visited.add(xref)
        if len(visited) > int(node_limit):
            raise RuntimeError("PDF 标签结构节点过多，已停止解析")

        object_value = xref_object(xref, compressed=False)
        if _value_has_table_structure_role(object_value, table_roles):
            return True

        stripped_object = str(object_value).lstrip()
        if stripped_object.startswith("["):
            queue.extend(
                int(value)
                for value in _PDF_XREF_REFERENCE.findall(str(object_value))
                if int(value) not in visited
            )
            continue

        role_kind, role_value = xref_get_key(xref, "S")
        if (
            role_kind == "name"
            and _decode_pdf_name(role_value) in table_roles
        ):
            return True

        child_kind, child_value = xref_get_key(xref, "K")
        if _value_has_table_structure_role(child_value, table_roles):
            return True
        if child_kind in {"xref", "array", "dict"}:
            queue.extend(
                int(value)
                for value in _PDF_XREF_REFERENCE.findall(str(child_value))
                if int(value) not in visited
            )
    return False

def _text_mapping_warning_count(text) -> int:
    count = 0
    for character in str(text or ""):
        codepoint = ord(character)
        category = unicodedata.category(character)
        if (
            character in {"\ufffd", "\ufffe", "\uffff"}
            or category in {"Cn", "Co", "Cs"}
            or 0xF0000 <= codepoint <= 0xFFFFD
            or 0x100000 <= codepoint <= 0x10FFFD
        ):
            count += 1
    return count


def _page_text_risks(page) -> tuple[int, int, int, str, int]:
    checkbox_count = 0
    symbol_font_runs = 0
    text_dictionary = page.get_text("dict")
    selectable_text_characters = 0
    selectable_text_parts = []
    mapping_warning_count = 0

    for block in text_dictionary.get("blocks", ()):
        for line in block.get("lines", ()):
            for span in line.get("spans", ()):
                text = span.get("text", "") or ""
                selectable_text_parts.append(text)
                checkbox_count += sum(
                    text.count(symbol) for symbol in _CHECKBOX_SYMBOLS
                )

                selectable_text_characters += sum(
                    not character.isspace() for character in text
                )
                mapping_warning_count += _text_mapping_warning_count(text)
                font_name = str(span.get("font", "")).casefold()
                if any(name in font_name for name in _SYMBOL_FONT_NAMES):
                    symbol_font_runs += 1

    return (
        checkbox_count,
        symbol_font_runs,
        selectable_text_characters,
        "".join(selectable_text_parts),
        mapping_warning_count,
    )


def _table_bbox_text(page, table) -> str:
    """Read a table text baseline, failing closed when it cannot be verified."""

    bbox = getattr(table, "bbox", None)
    if bbox is None:
        return ""
    try:
        return page.get_textbox(pymupdf.Rect(bbox)) or ""
    except Exception as exc:
        raise RuntimeError("无法读取 PDF 表格边界框内的文字") from exc


def _table_rect(table):
    bbox = getattr(table, "bbox", None)
    if bbox is None:
        return None
    try:
        rect = pymupdf.Rect(bbox)
    except Exception:
        return None
    if rect.is_empty or rect.width <= 0 or rect.height <= 0:
        return None
    return rect


def _table_layout_exclusion_rects(page, tables):
    """Return table bounds in the coordinates used by text and drawings."""

    rects = tuple(
        rect
        for table in tables
        if (rect := _table_rect(table)) is not None
    )
    rotation = getattr(page, "rotation", 0)
    if not isinstance(rotation, int) or rotation % 360 == 0:
        return rects

    # PyMuPDF table detection reports rotated-page coordinates, while text
    # words and drawing paths remain in unrotated page coordinates.
    try:
        derotation_matrix = page.derotation_matrix
        transformed = tuple(
            pymupdf.Rect(rect) * derotation_matrix
            for rect in rects
        )
    except Exception:
        # Keep the conservative fail-closed behavior if a backend cannot
        # expose or apply the page transform.
        return rects
    return tuple(
        rect
        for rect in transformed
        if not rect.is_empty and rect.width > 0 and rect.height > 0
    )


def _bbox_overlap_ratio(first, second) -> float:
    x0 = max(float(first.x0), float(second.x0))
    y0 = max(float(first.y0), float(second.y0))
    x1 = min(float(first.x1), float(second.x1))
    y1 = min(float(first.y1), float(second.y1))
    intersection = max(0.0, x1 - x0) * max(0.0, y1 - y0)
    smaller_area = min(float(first.width * first.height), float(second.width * second.height))
    return intersection / smaller_area if smaller_area > 0 else 0.0


def _bbox_nearly_same(first, second, tolerance: float = 2.0) -> bool:
    return all(
        abs(float(first[index]) - float(second[index])) <= float(tolerance)
        for index in range(4)
    )


def _table_candidate_text(table, extracted_rows_by_id=None) -> str:
    extracted_rows_by_id = extracted_rows_by_id or {}
    if id(table) in extracted_rows_by_id:
        rows = extracted_rows_by_id[id(table)]
        if rows is None:
            return ""
    else:
        try:
            rows = tuple(table.extract() or ())
        except Exception:
            return ""
    return "".join(
        character
        for row in rows
        for cell in (row or ())
        for character in str(cell or "")
        if not character.isspace()
    )


def _deduplicate_table_candidates(tables, extracted_rows_by_id=None):
    extracted_rows_by_id = extracted_rows_by_id or {}
    deduplicated = []
    for table in tables:
        if any(table is existing for existing in deduplicated):
            continue
        table_rect = _table_rect(table)
        overlapping = []
        if table_rect is not None:
            overlapping = [
                existing
                for existing in deduplicated
                if (
                    (existing_rect := _table_rect(existing)) is not None
                    and _bbox_overlap_ratio(table_rect, existing_rect) >= 0.7
                )
            ]
        # Text strategy can emit one artificial table spanning two already
        # detected tables. Reject only that union-like candidate; keep a single
        # genuinely nested table for independent validation.
        if len(overlapping) >= 2:
            continue
        if len(overlapping) == 1:
            existing = overlapping[0]
            existing_rect = _table_rect(existing)
            candidate_text = _table_candidate_text(
                table,
                extracted_rows_by_id,
            )
            same_text = bool(
                candidate_text
                and candidate_text
                == _table_candidate_text(existing, extracted_rows_by_id)
            )
            same_columns = int(getattr(table, "col_count", 0) or 0) == int(
                getattr(existing, "col_count", 0) or 0
            )
            if (
                existing_rect is not None
                and (
                    _bbox_nearly_same(table_rect, existing_rect)
                    or (same_text and same_columns)
                )
            ):
                continue
        deduplicated.append(table)

    def reading_order(table):
        rect = _table_rect(table)
        if rect is None:
            return (float("inf"), float("inf"), float("inf"), float("inf"))
        return (float(rect.y0), float(rect.x0), float(rect.y1), float(rect.x1))

    return tuple(sorted(deduplicated, key=reading_order))


def _normalized_table_row_texts(rows):
    return tuple(
        tuple(
            normalized
            for cell in (row or ())
            if (normalized := "".join(str(cell or "").split()))
        )
        for row in (rows or ())
    )


def _precise_grid_adds_unowned_narrow_columns(
    default_widths,
    precise_spans,
    precise_widths,
):
    """Detect narrow precision-only columns that never form a real cell."""

    default_widths = tuple(float(width) for width in default_widths)
    precise_widths = tuple(float(width) for width in precise_widths)
    if (
        len(precise_widths) <= len(default_widths)
        or not default_widths
        or any(width <= 0 for width in default_widths)
        or any(width <= 0 for width in precise_widths)
    ):
        return False

    default_total = sum(default_widths)
    precise_total = sum(precise_widths)
    if default_total <= 0 or precise_total <= 0:
        return False

    scale = precise_total / default_total
    default_boundaries = []
    position = 0.0
    for width in default_widths[:-1]:
        position += width
        default_boundaries.append(position * scale)

    precise_boundaries = [0.0]
    for width in precise_widths:
        precise_boundaries.append(precise_boundaries[-1] + width)

    boundary_tolerance = max(0.5, precise_total * 0.0005)
    narrow_limit = max(
        4.0,
        float(_median_coordinate(precise_widths) or 0.0) * 0.2,
    )
    unit_owned_columns = {
        int(column_index)
        for _row_index, column_index, _row_span, column_span in precise_spans
        if int(column_span) == 1
    }

    unmatched_precise_boundaries = {
        index
        for index, boundary in enumerate(precise_boundaries[1:-1], start=1)
        if not any(
            abs(boundary - default_boundary) <= boundary_tolerance
            for default_boundary in default_boundaries
        )
    }
    for column_index, width in enumerate(precise_widths):
        if column_index in unit_owned_columns or width > narrow_limit:
            continue
        left = precise_boundaries[column_index]
        right = precise_boundaries[column_index + 1]
        interleaves_default_boundary = any(
            left + boundary_tolerance
            < default_boundary
            < right - boundary_tolerance
            for default_boundary in default_boundaries
        )
        touches_precision_only_boundary = (
            column_index in unmatched_precise_boundaries
            or column_index + 1 in unmatched_precise_boundaries
        )
        if interleaves_default_boundary or touches_precision_only_boundary:
            return True
    return False


def _select_precise_line_tables(
    default_tables,
    precise_tables,
    extracted_rows_by_id,
):
    """Prefer a stable 1 pt grid when the default 3 pt snap aliases boundaries."""

    selected = []
    used_precise_ids = set()
    matched_precise_ids = set()
    for default in default_tables:
        default_rect = _table_rect(default)
        default_rows = extracted_rows_by_id.get(id(default))
        default_shape = _effective_table_shape(default, default_rows or ())
        default_cell_count = _effective_table_cell_count(
            default,
            default_rows or (),
        )
        default_text_rows = _normalized_table_row_texts(default_rows)
        default_geometry = _table_geometry_model(default, default_shape)
        best = default
        best_columns = default_shape[1]

        if default_rect is not None and default_rows is not None:
            for precise in precise_tables:
                if id(precise) in used_precise_ids:
                    continue
                precise_rect = _table_rect(precise)
                precise_rows = extracted_rows_by_id.get(id(precise))
                if precise_rect is None or precise_rows is None:
                    continue
                if _bbox_nearly_same(default_rect, precise_rect):
                    matched_precise_ids.add(id(precise))
                precise_shape = _effective_table_shape(precise, precise_rows)
                if (
                    not _bbox_nearly_same(default_rect, precise_rect)
                    or precise_shape[0] != default_shape[0]
                    or precise_shape[1] <= best_columns
                    or _effective_table_cell_count(precise, precise_rows)
                    != default_cell_count
                    or _normalized_table_row_texts(precise_rows)
                    != default_text_rows
                ):
                    continue
                geometry = _table_geometry_model(precise, precise_shape)
                if geometry is None:
                    continue
                precise_spans, precise_widths = geometry
                if (
                    len(precise_spans) != default_cell_count
                    or len(precise_widths) != precise_shape[1]
                    or any(width <= 0 for width in precise_widths)
                ):
                    continue
                if default_geometry is not None:
                    _default_spans, default_widths = default_geometry
                    if _precise_grid_adds_unowned_narrow_columns(
                        default_widths,
                        precise_spans,
                        precise_widths,
                    ):
                        continue
                best = precise
                best_columns = precise_shape[1]

        selected.append(best)
        if best is not default:
            used_precise_ids.add(id(best))
    selected.extend(
        precise
        for precise in precise_tables
        if id(precise) not in matched_precise_ids
        and id(precise) not in used_precise_ids
    )
    return _deduplicate_table_candidates(
        tuple(selected),
        extracted_rows_by_id=extracted_rows_by_id,
    )


def _table_exclusion_clips(page, tables):
    try:
        page_rect = pymupdf.Rect(page.rect)
    except Exception:
        return ()

    clips = []
    seen = set()
    for table in tables:
        table_rect = _table_rect(table)
        if table_rect is None:
            continue
        candidates = (
            pymupdf.Rect(page_rect.x0, page_rect.y0, page_rect.x1, table_rect.y0),
            pymupdf.Rect(page_rect.x0, table_rect.y1, page_rect.x1, page_rect.y1),
            pymupdf.Rect(page_rect.x0, table_rect.y0, table_rect.x0, table_rect.y1),
            pymupdf.Rect(table_rect.x1, table_rect.y0, page_rect.x1, table_rect.y1),
        )
        for clip in candidates:
            if clip.is_empty or clip.width < 8 or clip.height < 8:
                continue
            key = tuple(round(float(value), 3) for value in clip)
            if key in seen:
                continue
            seen.add(key)
            clips.append(clip)
    return tuple(clips)


def _has_repeated_large_column_gaps(page, table) -> bool:
    """Reject prose falsely segmented into columns by the text table strategy."""

    bbox = getattr(table, "bbox", None)
    if bbox is None:
        return False
    try:
        table_rect = pymupdf.Rect(bbox)
        words = tuple(page.get_text("words") or ())
    except Exception as exc:
        raise RuntimeError("无法验证 PDF 表格候选的文字列") from exc

    inside_words = []
    for word in words:
        if len(word) < 5:
            continue
        x0, y0, x1, y1 = map(float, word[:4])
        center_x = (x0 + x1) / 2
        center_y = (y0 + y1) / 2
        if (
            table_rect.x0 <= center_x <= table_rect.x1
            and table_rect.y0 <= center_y <= table_rect.y1
        ):
            inside_words.append((x0, y0, x1, y1, center_y))
    if len(inside_words) < 4:
        return False

    rows: list[dict[str, object]] = []
    for word in sorted(inside_words, key=lambda item: (item[4], item[0])):
        height = max(1.0, word[3] - word[1])
        if rows and abs(word[4] - float(rows[-1]["center"])) <= max(
            2.5,
            height * 0.45,
        ):
            row_words = rows[-1]["words"]
            row_words.append(word)
            rows[-1]["center"] = sum(item[4] for item in row_words) / len(row_words)
        else:
            rows.append({"center": word[4], "words": [word]})

    column_count = int(getattr(table, "col_count", 0) or 0)
    required_gaps = min(2, max(1, column_count - 1))
    qualifying_rows = 0
    for row in rows:
        row_words = sorted(row["words"], key=lambda item: item[0])
        heights = sorted(max(1.0, word[3] - word[1]) for word in row_words)
        median_height = heights[len(heights) // 2]
        gap_threshold = max(8.0, median_height * 0.65)
        large_gaps = sum(
            next_word[0] - word[2] >= gap_threshold
            for word, next_word in zip(row_words, row_words[1:])
        )
        if large_gaps >= required_gaps:
            qualifying_rows += 1

    return qualifying_rows >= 2


def _has_strict_single_row_columns(page, table) -> bool:
    if (
        int(getattr(table, "row_count", 0) or 0) != 1
        or int(getattr(table, "col_count", 0) or 0) < 3
    ):
        return False

    table_rect = _table_rect(table)
    if table_rect is None:
        return False
    try:
        extracted_rows = tuple(table.extract() or ())
    except Exception as exc:
        raise RuntimeError("无法验证 PDF 单行表格候选") from exc
    if len(extracted_rows) != 1:
        return False

    cell_texts = [
        str(cell or "").strip()
        for cell in (extracted_rows[0] or ())
        if str(cell or "").strip()
    ]
    if len(cell_texts) < 3:
        return False
    if any(":" in text or chr(0xFF1A) in text for text in cell_texts):
        return False
    try:
        words = tuple(page.get_text("words") or ())
    except Exception as exc:
        raise RuntimeError("无法读取 PDF 单行表格候选的文字位置") from exc

    inside_words = []
    for word in words:
        if len(word) < 5:
            continue
        x0, y0, x1, y1 = map(float, word[:4])
        center_x = (x0 + x1) / 2
        center_y = (y0 + y1) / 2
        if (
            table_rect.x0 <= center_x <= table_rect.x1
            and table_rect.y0 <= center_y <= table_rect.y1
        ):
            inside_words.append((x0, y0, x1, y1))
    if len(inside_words) < 3:
        return False

    inside_words.sort(key=lambda item: item[0])
    heights = sorted(max(1.0, word[3] - word[1]) for word in inside_words)
    median_height = heights[len(heights) // 2]
    gap_threshold = max(24.0, median_height * 2.0)
    large_gaps = sum(
        next_word[0] - word[2] >= gap_threshold
        for word, next_word in zip(inside_words, inside_words[1:])
    )
    return large_gaps >= 2


def _find_text_table_candidates(
    page,
    find_tables,
    clips=(),
    extracted_rows_by_id=None,
    cancel_requested=None,
):
    extracted_rows_by_id = (
        extracted_rows_by_id if extracted_rows_by_id is not None else {}
    )
    candidates = []
    search_clips = tuple(clips) if clips else (None,)
    for clip in search_clips:
        _raise_if_pdf_analysis_cancelled(cancel_requested)
        clip_argument = {} if clip is None else {"clip": clip}
        multi_finder = find_tables(
            strategy="text",
            min_words_vertical=2,
            **clip_argument,
        )
        _raise_if_pdf_analysis_cancelled(cancel_requested)
        multi_tables = tuple(getattr(multi_finder, "tables", ()) or ())
        for table in multi_tables:
            try:
                extracted_rows_by_id[id(table)] = tuple(table.extract() or ())
            except Exception:
                extracted_rows_by_id[id(table)] = None
        candidates.extend(
            table
            for table in multi_tables
            if int(getattr(table, "row_count", 0) or 0) >= 2
            and int(getattr(table, "col_count", 0) or 0) >= 2
            and _has_repeated_large_column_gaps(page, table)
        )

        _raise_if_pdf_analysis_cancelled(cancel_requested)
        single_finder = find_tables(
            strategy="text",
            min_words_vertical=1,
            **clip_argument,
        )
        _raise_if_pdf_analysis_cancelled(cancel_requested)
        single_tables = tuple(getattr(single_finder, "tables", ()) or ())
        for table in single_tables:
            try:
                extracted_rows_by_id[id(table)] = tuple(table.extract() or ())
            except Exception:
                extracted_rows_by_id[id(table)] = None
        candidates.extend(
            table
            for table in single_tables
            if _has_strict_single_row_columns(page, table)
        )
    return tuple(candidates)


def _effective_table_cell_count(
    table,
    extracted_rows,
    normalize_text_strategy_rows=False,
) -> int:
    rows = tuple(extracted_rows or ())
    if normalize_text_strategy_rows:
        nonempty_rows = tuple(
            row
            for row in rows
            if any(str(cell or "").strip() for cell in (row or ()))
        )
        if nonempty_rows:
            return sum(
                cell is not None
                for row in nonempty_rows
                for cell in (row or ())
            )
    try:
        table_cells = tuple(getattr(table, "cells", ()) or ())
    except TypeError:
        table_cells = ()
    physical_cell_count = sum(cell is not None for cell in table_cells)
    if physical_cell_count:
        return physical_cell_count
    return sum(
        cell is not None
        for row in rows
        for cell in (row or ())
    )

def _effective_table_shape(
    table,
    extracted_rows,
    normalize_text_strategy_rows=False,
) -> tuple[int, int]:
    detector_rows = max(0, int(getattr(table, "row_count", 0) or 0))
    detector_columns = max(0, int(getattr(table, "col_count", 0) or 0))
    rows = tuple(extracted_rows or ())
    extracted_columns = max((len(row or ()) for row in rows), default=0)
    if normalize_text_strategy_rows:
        nonempty_rows = tuple(
            row
            for row in rows
            if any(str(cell or "").strip() for cell in (row or ()))
        )
        if nonempty_rows:
            return (
                len(nonempty_rows),
                max(detector_columns, extracted_columns),
            )
    return (
        max(detector_rows, len(rows)),
        max(detector_columns, extracted_columns),
    )


def _median_coordinate(values):
    ordered = sorted(float(value) for value in values)
    if not ordered:
        return None
    midpoint = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[midpoint]
    return (ordered[midpoint - 1] + ordered[midpoint]) / 2.0


def _cell_spans_from_matrix(matrix, shape):
    try:
        expected_rows, expected_columns = (int(shape[0]), int(shape[1]))
        rows = tuple(tuple(row) for row in matrix)
    except (TypeError, ValueError, IndexError):
        return None
    if (
        expected_rows < 1
        or expected_columns < 1
        or len(rows) != expected_rows
        or any(len(row) != expected_columns for row in rows)
    ):
        return None
    spans = []
    for row_index, row in enumerate(rows):
        starts = [
            column_index
            for column_index, value in enumerate(row)
            if value is not None
        ]
        if not starts:
            return None
        for start_index, column_index in enumerate(starts):
            next_column = (
                starts[start_index + 1]
                if start_index + 1 < len(starts)
                else expected_columns
            )
            spans.append(
                (row_index, column_index, 1, next_column - column_index)
            )
    return tuple(spans)


def _table_geometry_model(table, shape):
    try:
        expected_rows, expected_columns = (int(shape[0]), int(shape[1]))
        rows = tuple(getattr(table, "rows", ()) or ())
        table_rect = pymupdf.Rect(getattr(table, "bbox"))
    except Exception:
        return None
    if (
        expected_rows < 1
        or expected_columns < 1
        or len(rows) != expected_rows
        or table_rect.is_empty
    ):
        return None

    x_candidates = {index: [] for index in range(expected_columns)}
    y_candidates = {index: [] for index in range(expected_rows)}
    row_cells = []
    for row_index, row in enumerate(rows):
        try:
            row_rect = pymupdf.Rect(getattr(row, "bbox"))
            cells = tuple(getattr(row, "cells", ()) or ())
        except (TypeError, ValueError):
            return None
        if len(cells) != expected_columns or row_rect.is_empty:
            return None
        y_candidates[row_index].append(float(row_rect.y0))
        normalized_cells = []
        for column_index, cell in enumerate(cells):
            if cell is None:
                normalized_cells.append(None)
                continue
            try:
                cell_rect = pymupdf.Rect(cell)
            except (TypeError, ValueError):
                return None
            if cell_rect.is_empty:
                return None
            x_candidates[column_index].append(float(cell_rect.x0))
            y_candidates[row_index].append(float(cell_rect.y0))
            normalized_cells.append(cell_rect)
        row_cells.append(tuple(normalized_cells))

    x_boundaries = []
    for column_index in range(expected_columns):
        coordinate = _median_coordinate(x_candidates[column_index])
        if coordinate is None:
            return None
        x_boundaries.append(coordinate)
    x_boundaries.append(float(table_rect.x1))
    y_boundaries = []
    for row_index in range(expected_rows):
        coordinate = _median_coordinate(y_candidates[row_index])
        if coordinate is None:
            return None
        y_boundaries.append(coordinate)
    y_boundaries.append(float(table_rect.y1))

    if any(
        next_value <= value
        for value, next_value in zip(x_boundaries, x_boundaries[1:])
    ) or any(
        next_value <= value
        for value, next_value in zip(y_boundaries, y_boundaries[1:])
    ):
        return None

    tolerance = max(
        0.5,
        float(table_rect.width) * 0.0001,
        float(table_rect.height) * 0.0001,
    )

    def boundary_index(boundaries, coordinate, minimum_index):
        matches = [
            (abs(value - float(coordinate)), index)
            for index, value in enumerate(boundaries)
            if index >= minimum_index
            and abs(value - float(coordinate)) <= tolerance
        ]
        return min(matches)[1] if matches else None

    coverage = [
        [False] * expected_columns for _row_index in range(expected_rows)
    ]
    spans = []
    for row_index, cells in enumerate(row_cells):
        for column_index, cell_rect in enumerate(cells):
            if cell_rect is None:
                continue
            if (
                abs(float(cell_rect.x0) - x_boundaries[column_index])
                > tolerance
                or abs(float(cell_rect.y0) - y_boundaries[row_index])
                > tolerance
            ):
                return None
            column_end = boundary_index(
                x_boundaries,
                cell_rect.x1,
                column_index + 1,
            )
            row_end = boundary_index(
                y_boundaries,
                cell_rect.y1,
                row_index + 1,
            )
            if column_end is None or row_end is None:
                return None
            row_span = row_end - row_index
            column_span = column_end - column_index
            if row_span < 1 or column_span < 1:
                return None
            for covered_row in range(row_index, row_end):
                for covered_column in range(column_index, column_end):
                    if coverage[covered_row][covered_column]:
                        return None
                    coverage[covered_row][covered_column] = True
            spans.append(
                (row_index, column_index, row_span, column_span)
            )
    if not all(all(row) for row in coverage):
        return None
    column_widths = tuple(
        next_value - value
        for value, next_value in zip(x_boundaries, x_boundaries[1:])
    )
    return tuple(sorted(spans)), column_widths


def _table_cell_spans_from_geometry(table, shape):
    model = _table_geometry_model(table, shape)
    return None if model is None else model[0]


def _table_column_widths_from_geometry(table, shape):
    model = _table_geometry_model(table, shape)
    return None if model is None else model[1]


def _table_full_width_span_flags(table, retained_row_indices, expected_rows):
    table_rect = _table_rect(table)
    flags = [False] * max(0, int(expected_rows))
    if table_rect is None:
        return tuple(flags)
    try:
        table_rows = tuple(getattr(table, "rows", ()) or ())
    except TypeError:
        return tuple(flags)

    tolerance = max(1.0, float(table_rect.width) * 0.002)
    for output_index, source_index in enumerate(retained_row_indices):
        if output_index >= len(flags) or not 0 <= source_index < len(table_rows):
            continue
        try:
            row_cells = tuple(getattr(table_rows[source_index], "cells", ()) or ())
            present_cells = [
                pymupdf.Rect(cell) for cell in row_cells if cell is not None
            ]
        except Exception:
            continue
        if len(present_cells) != 1:
            continue
        cell_rect = present_cells[0]
        flags[output_index] = bool(
            abs(float(cell_rect.x0) - float(table_rect.x0)) <= tolerance
            and abs(float(cell_rect.x1) - float(table_rect.x1)) <= tolerance
        )
    return tuple(flags)


def _page_word_rows(page, excluded_rects=()):
    try:
        raw_words = tuple(page.get_text("words") or ())
    except Exception as exc:
        raise RuntimeError("无法读取 PDF 文字坐标以检查表格版面") from exc

    words = []
    for word in raw_words:
        if len(word) < 5:
            continue
        text = str(word[4] or "").strip()
        if not text:
            continue
        try:
            x0, y0, x1, y1 = map(float, word[:4])
        except (TypeError, ValueError):
            continue
        if x1 <= x0 or y1 <= y0:
            continue
        center_x = (x0 + x1) / 2
        center_y = (y0 + y1) / 2
        if any(
            rect.x0 - 2 <= center_x <= rect.x1 + 2
            and rect.y0 - 2 <= center_y <= rect.y1 + 2
            for rect in excluded_rects
        ):
            continue
        words.append((x0, y0, x1, y1, text, center_y))

    rows: list[dict[str, object]] = []
    for word in sorted(words, key=lambda item: (item[5], item[0])):
        height = max(1.0, word[3] - word[1])
        if rows and abs(word[5] - float(rows[-1]["center"])) <= max(
            2.5,
            height * 0.45,
        ):
            row_words = rows[-1]["words"]
            row_words.append(word)
            rows[-1]["center"] = sum(item[5] for item in row_words) / len(row_words)
        else:
            rows.append({"center": word[5], "words": [word]})
    return tuple(tuple(row["words"]) for row in rows)


def _row_column_chunks(row_words, minimum_gap=10.0, height_ratio=0.9):
    words = sorted(row_words, key=lambda item: item[0])
    if not words:
        return ()
    heights = sorted(max(1.0, word[3] - word[1]) for word in words)
    median_height = heights[len(heights) // 2]
    gap_threshold = max(float(minimum_gap), median_height * float(height_ratio))
    chunks = [[words[0]]]
    for previous, current in zip(words, words[1:]):
        if current[0] - previous[2] >= gap_threshold:
            chunks.append([current])
        else:
            chunks[-1].append(current)
    return tuple(
        (
            min(word[0] for word in chunk),
            max(word[2] for word in chunk),
            " ".join(word[4] for word in chunk),
        )
        for chunk in chunks
    )


def _aligned_column_count(first_chunks, second_chunks, tolerance=10.0):
    remaining = [chunk[0] for chunk in second_chunks]
    aligned = 0
    for first_anchor in (chunk[0] for chunk in first_chunks):
        if not remaining:
            break
        best_index = min(
            range(len(remaining)),
            key=lambda index: abs(remaining[index] - first_anchor),
        )
        if abs(remaining[best_index] - first_anchor) <= tolerance:
            aligned += 1
            remaining.pop(best_index)
    return aligned


def _has_repeated_text_columns(
    page,
    excluded_rects=(),
    allow_single_row=True,
) -> bool:
    rows = _page_word_rows(page, excluded_rects=excluded_rects)
    chunked_rows = [
        (row, _row_column_chunks(row))
        for row in rows
    ]
    multi_column_rows = [
        (row, chunks)
        for row, chunks in chunked_rows
        if len(chunks) >= 2
    ]
    for first_index, (_first_row, first_chunks) in enumerate(multi_column_rows):
        for _second_row, second_chunks in multi_column_rows[first_index + 1 :]:
            if _aligned_column_count(first_chunks, second_chunks) >= 2:
                return True

    # Very compact borderless cells can be separated by less than a normal
    # line-height. Use a stricter repeated-anchor check so a 2x2 table with a
    # narrow gutter is caught without treating ordinary prose spaces as cells.
    compact_rows = []
    for row_index, row in enumerate(rows):
        chunks = _row_column_chunks(
            row,
            minimum_gap=3.0,
            height_ratio=0.2,
        )
        if len(chunks) >= 2:
            compact_rows.append((row_index, row, chunks))

    for first_index, (_row_index, first_row, first_chunks) in enumerate(compact_rows):
        first_anchors = tuple(chunk[0] for chunk in first_chunks)
        first_center = sum(word[5] for word in first_row) / len(first_row)
        first_height = max(word[3] - word[1] for word in first_row)
        for _other_index, second_row, second_chunks in compact_rows[
            first_index + 1 :
        ]:
            if len(first_chunks) != len(second_chunks):
                continue
            second_center = sum(word[5] for word in second_row) / len(second_row)
            second_height = max(word[3] - word[1] for word in second_row)
            minimum_row_separation = max(
                4.0,
                max(first_height, second_height) * 0.8,
            )
            if abs(second_center - first_center) < minimum_row_separation:
                continue
            second_anchors = tuple(chunk[0] for chunk in second_chunks)
            if all(
                abs(first_anchor - second_anchor) <= 2.0
                for first_anchor, second_anchor in zip(
                    first_anchors,
                    second_anchors,
                )
            ):
                return True

    if not allow_single_row:
        return False

    # A compact data row may sit below a merged heading, so the neighboring
    # rows do not necessarily expose the same number of column chunks.
    for row_index, row, chunks in compact_rows:
        if len(chunks) > 4:
            continue
        row_height = max(
            max(1.0, word[3] - word[1])
            for word in row
        )
        row_center = sum(word[5] for word in row) / len(row)
        table_x0 = min(chunk[0] for chunk in chunks)
        table_x1 = max(chunk[1] for chunk in chunks)
        for neighbor_index in range(
            max(0, row_index - 2),
            min(len(rows), row_index + 3),
        ):
            if neighbor_index == row_index:
                continue
            neighbor = rows[neighbor_index]
            neighbor_center = sum(word[5] for word in neighbor) / len(neighbor)
            if abs(neighbor_center - row_center) > max(30.0, row_height * 4):
                continue
            neighbor_x0 = min(word[0] for word in neighbor)
            neighbor_x1 = max(word[2] for word in neighbor)
            if (
                neighbor_x1 >= table_x0 - 2
                and neighbor_x0 <= table_x1 + 2
            ):
                return True

    for row, _chunks in chunked_rows:
        conservative_chunks = _row_column_chunks(
            row,
            minimum_gap=8.0,
            height_ratio=0.6,
        )
        if len(conservative_chunks) < 2:
            continue
        # A single row of separated text blocks is indistinguishable from a
        # one-row borderless table in PDF drawing coordinates. Fail closed so
        # a narrow two-column table or a label containing a colon cannot bypass
        # the editable-table output gate.
        return True
    return False


def _small_vector_mark_count(drawings) -> int:
    """Count compact diagonal paths that may encode a drawn check or cross."""

    count = 0
    for drawing in drawings:
        path_points = []
        diagonal_found = False
        for item in drawing.get("items", ()) or ():
            if not item:
                continue
            try:
                if item[0] == "l" and len(item) >= 3:
                    points = (item[1], item[2])
                elif item[0] == "c" and len(item) >= 5:
                    points = tuple(item[index] for index in range(1, 5))
                else:
                    continue
                coordinates = tuple(
                    (float(point.x), float(point.y))
                    for point in points
                )
            except (AttributeError, TypeError, ValueError):
                continue

            path_points.extend(coordinates)
            for first, second in zip(coordinates, coordinates[1:]):
                delta_x = abs(second[0] - first[0])
                delta_y = abs(second[1] - first[1])
                if delta_x >= 1.5 and delta_y >= 1.5:
                    diagonal_found = True

        if not diagonal_found or not path_points:
            continue
        x_values = tuple(point[0] for point in path_points)
        y_values = tuple(point[1] for point in path_points)
        width = max(x_values) - min(x_values)
        height = max(y_values) - min(y_values)
        if 1.5 <= width <= 48 and 1.5 <= height <= 48:
            count += 1
    return count

_TABLE_BORDER_MATCH_TOLERANCE = 3.0
_TABLE_BORDER_AXIS_TOLERANCE = 1.0
_TABLE_BORDER_GAP_TOLERANCE = 1.0
_TABLE_BORDER_MINIMUM_COVERAGE = 0.7
_TABLE_BORDER_FILLED_RULE_MAX_THICKNESS = 3.0
_TABLE_BORDER_FILLED_RULE_MIN_ASPECT_RATIO = 4.0


def _drawing_has_visible_stroke(drawing) -> bool:
    try:
        drawing_type = str(drawing.get("type", "")).strip().casefold()
        color = drawing.get("color")
        opacity = float(drawing.get("stroke_opacity", 1.0))
        width = drawing.get("width", 1.0)
        width = 1.0 if width is None else float(width)
    except (AttributeError, TypeError, ValueError):
        return False
    if drawing_type and "s" not in drawing_type:
        return False
    return color is not None and opacity > 0 and width >= 0


def _drawing_has_visible_fill(drawing) -> bool:
    try:
        drawing_type = str(drawing.get("type", "")).strip().casefold()
        color = drawing.get("fill")
        opacity = float(drawing.get("fill_opacity", 1.0))
    except (AttributeError, TypeError, ValueError):
        return False
    if drawing_type and "f" not in drawing_type:
        return False
    return color is not None and opacity > 0


def _point_coordinates(point):
    try:
        return float(point.x), float(point.y)
    except (AttributeError, TypeError, ValueError):
        return None


def _drawing_item_segments(item):
    if not item:
        return ()
    kind = item[0]
    if kind == "l" and len(item) >= 3:
        first = _point_coordinates(item[1])
        second = _point_coordinates(item[2])
        return ((first, second),) if first and second else ()
    if kind == "re" and len(item) >= 2:
        try:
            rect = pymupdf.Rect(item[1])
        except Exception:
            return ()
        if rect.is_empty:
            return ()
        top_left = (float(rect.x0), float(rect.y0))
        top_right = (float(rect.x1), float(rect.y0))
        bottom_right = (float(rect.x1), float(rect.y1))
        bottom_left = (float(rect.x0), float(rect.y1))
        return (
            (top_left, top_right),
            (top_right, bottom_right),
            (bottom_right, bottom_left),
            (bottom_left, top_left),
        )
    if kind == "qu" and len(item) >= 2:
        try:
            quad = item[1]
            points = tuple(
                _point_coordinates(point)
                for point in (quad.ul, quad.ur, quad.lr, quad.ll)
            )
        except Exception:
            return ()
        if any(point is None for point in points):
            return ()
        return tuple(
            (first, second)
            for first, second in zip(points, points[1:] + points[:1])
        )
    if kind == "c" and len(item) >= 5:
        points = tuple(_point_coordinates(item[index]) for index in range(1, 5))
        if any(point is None for point in points):
            return ()
        x_values = tuple(point[0] for point in points)
        y_values = tuple(point[1] for point in points)
        if (
            max(x_values) - min(x_values) <= _TABLE_BORDER_AXIS_TOLERANCE
            or max(y_values) - min(y_values) <= _TABLE_BORDER_AXIS_TOLERANCE
        ):
            return ((points[0], points[-1]),)
    return ()


def _filled_rectangle_centerline(item):
    if not item or item[0] != "re" or len(item) < 2:
        return ()
    try:
        rect = pymupdf.Rect(item[1])
    except Exception:
        return ()
    if rect.is_empty:
        return ()
    width = float(rect.width)
    height = float(rect.height)
    thickness = min(width, height)
    length = max(width, height)
    if (
        thickness <= 0
        or thickness > _TABLE_BORDER_FILLED_RULE_MAX_THICKNESS
        or length / thickness < _TABLE_BORDER_FILLED_RULE_MIN_ASPECT_RATIO
    ):
        return ()
    if width >= height:
        center_y = (float(rect.y0) + float(rect.y1)) / 2
        return (((float(rect.x0), center_y), (float(rect.x1), center_y)),)
    center_x = (float(rect.x0) + float(rect.x1)) / 2
    return (((center_x, float(rect.y0)), (center_x, float(rect.y1))),)


def _visible_drawing_line_segments(drawings):
    segments = []
    for drawing in drawings:
        visible_stroke = _drawing_has_visible_stroke(drawing)
        visible_fill = _drawing_has_visible_fill(drawing)
        for item in drawing.get("items", ()) or ():
            if visible_stroke:
                for first, second in _drawing_item_segments(item):
                    if first != second:
                        segments.append((first, second))
            elif visible_fill:
                segments.extend(_filled_rectangle_centerline(item))
    return tuple(segments)


def _table_edge_in_drawing_space(page, first, second):
    try:
        first_point = pymupdf.Point(*first)
        second_point = pymupdf.Point(*second)
    except Exception:
        return None
    rotation = getattr(page, "rotation", 0)
    if isinstance(rotation, int) and rotation % 360:
        try:
            matrix = page.derotation_matrix
            first_point = first_point * matrix
            second_point = second_point * matrix
        except Exception:
            return None
    return (
        (float(first_point.x), float(first_point.y)),
        (float(second_point.x), float(second_point.y)),
    )


def _line_segments_cover_edge(
    first,
    second,
    segments,
    perpendicular_extent,
) -> bool:
    delta_x = abs(second[0] - first[0])
    delta_y = abs(second[1] - first[1])
    edge_length = max(delta_x, delta_y)
    if edge_length <= 0:
        return False
    horizontal = delta_x >= delta_y
    coordinate_tolerance = min(
        _TABLE_BORDER_MATCH_TOLERANCE,
        max(0.1, float(perpendicular_extent) * 0.45),
    )
    endpoint_tolerance = min(
        _TABLE_BORDER_MATCH_TOLERANCE,
        max(0.1, edge_length * 0.45),
    )
    target_coordinate = (
        (first[1] + second[1]) / 2
        if horizontal
        else (first[0] + second[0]) / 2
    )
    target_start, target_end = sorted(
        (first[0], second[0]) if horizontal else (first[1], second[1])
    )

    intervals = []
    for segment_first, segment_second in segments:
        segment_delta_x = abs(segment_second[0] - segment_first[0])
        segment_delta_y = abs(segment_second[1] - segment_first[1])
        if horizontal:
            axis_limit = min(
                _TABLE_BORDER_AXIS_TOLERANCE,
                max(0.1, segment_delta_x * 0.05),
            )
            if segment_delta_x <= 0 or segment_delta_y > axis_limit:
                continue
            if max(
                abs(segment_first[1] - target_coordinate),
                abs(segment_second[1] - target_coordinate),
            ) > coordinate_tolerance:
                continue
            segment_start, segment_end = sorted(
                (segment_first[0], segment_second[0])
            )
        else:
            axis_limit = min(
                _TABLE_BORDER_AXIS_TOLERANCE,
                max(0.1, segment_delta_y * 0.05),
            )
            if segment_delta_y <= 0 or segment_delta_x > axis_limit:
                continue
            if max(
                abs(segment_first[0] - target_coordinate),
                abs(segment_second[0] - target_coordinate),
            ) > coordinate_tolerance:
                continue
            segment_start, segment_end = sorted(
                (segment_first[1], segment_second[1])
            )
        if (
            segment_end < target_start - endpoint_tolerance
            or segment_start > target_end + endpoint_tolerance
        ):
            continue
        clipped_start = max(target_start, segment_start)
        clipped_end = min(target_end, segment_end)
        if clipped_end > clipped_start:
            intervals.append((clipped_start, clipped_end))

    if not intervals:
        return False
    intervals.sort()
    covered_length = 0.0
    covered_until = target_start
    for interval_start, interval_end in intervals:
        if interval_end <= covered_until:
            continue
        covered_length += interval_end - max(interval_start, covered_until)
        covered_until = interval_end
    if (
        covered_length + 1e-9
        < edge_length * _TABLE_BORDER_MINIMUM_COVERAGE
    ):
        return False
    cursor = target_start
    started = False
    for interval_start, interval_end in intervals:
        if not started:
            if interval_start > target_start + endpoint_tolerance:
                return False
            cursor = max(cursor, interval_end)
            started = True
            continue
        if interval_start > cursor + _TABLE_BORDER_GAP_TOLERANCE:
            return False
        cursor = max(cursor, interval_end)
        if cursor >= target_end - endpoint_tolerance:
            return True
    return started and cursor >= target_end - endpoint_tolerance


def _table_cell_border_edges_from_drawings(page, table, spans, drawings):
    """Return explicit PDF stroke evidence in (top, left, bottom, right) order."""

    spans = tuple(sorted(tuple(span) for span in (spans or ())))
    if not spans:
        return ()
    segments = _visible_drawing_line_segments(drawings)
    try:
        rows = tuple(getattr(table, "rows", ()) or ())
    except TypeError:
        rows = ()

    results = []
    for row_index, column_index, _row_span, _column_span in spans:
        try:
            cells = tuple(getattr(rows[row_index], "cells", ()) or ())
            cell_value = cells[column_index]
            if cell_value is None:
                raise ValueError("missing owner cell rectangle")
            rect = pymupdf.Rect(cell_value)
            if rect.is_empty:
                raise ValueError("empty owner cell rectangle")
        except (AttributeError, IndexError, TypeError, ValueError):
            results.append((False, False, False, False))
            continue

        edge_values = (
            ((rect.x0, rect.y0), (rect.x1, rect.y0), rect.height),
            ((rect.x0, rect.y0), (rect.x0, rect.y1), rect.width),
            ((rect.x0, rect.y1), (rect.x1, rect.y1), rect.height),
            ((rect.x1, rect.y0), (rect.x1, rect.y1), rect.width),
        )
        edges = []
        for first, second, perpendicular_extent in edge_values:
            transformed = _table_edge_in_drawing_space(page, first, second)
            edges.append(
                bool(
                    transformed
                    and _line_segments_cover_edge(
                        transformed[0],
                        transformed[1],
                        segments,
                        perpendicular_extent,
                    )
                )
            )
        results.append(tuple(edges))
    return tuple(results)

def _drawing_axis_segments(drawings, excluded_rects=()):
    horizontal = []
    vertical = []

    def add_segment(x0, y0, x1, y1):
        center_x = (x0 + x1) / 2
        center_y = (y0 + y1) / 2
        if any(
            rect.x0 - 2 <= center_x <= rect.x1 + 2
            and rect.y0 - 2 <= center_y <= rect.y1 + 2
            for rect in excluded_rects
        ):
            return
        if abs(y1 - y0) <= 1.5 and abs(x1 - x0) >= 12:
            horizontal.append((min(x0, x1), max(x0, x1), (y0 + y1) / 2))
        elif abs(x1 - x0) <= 1.5 and abs(y1 - y0) >= 12:
            vertical.append(((x0 + x1) / 2, min(y0, y1), max(y0, y1)))

    for drawing in drawings:
        for item in drawing.get("items", ()) or ():
            if not item:
                continue
            if item[0] == "l" and len(item) >= 3:
                try:
                    first, second = item[1], item[2]
                    add_segment(
                        float(first.x),
                        float(first.y),
                        float(second.x),
                        float(second.y),
                    )
                except (AttributeError, TypeError, ValueError):
                    continue
            elif item[0] == "c" and len(item) >= 5:
                try:
                    points = tuple(item[index] for index in range(1, 5))
                    x_values = tuple(float(point.x) for point in points)
                    y_values = tuple(float(point.y) for point in points)
                    if (
                        max(y_values) - min(y_values) <= 1.5
                        or max(x_values) - min(x_values) <= 1.5
                    ):
                        add_segment(
                            x_values[0],
                            y_values[0],
                            x_values[-1],
                            y_values[-1],
                        )
                except (AttributeError, TypeError, ValueError):
                    continue
            elif item[0] == "re" and len(item) >= 2:
                try:
                    rect = pymupdf.Rect(item[1])
                except Exception:
                    continue
                add_segment(rect.x0, rect.y0, rect.x1, rect.y0)
                add_segment(rect.x0, rect.y1, rect.x1, rect.y1)
                add_segment(rect.x0, rect.y0, rect.x0, rect.y1)
                add_segment(rect.x1, rect.y0, rect.x1, rect.y1)
            elif item[0] == "qu" and len(item) >= 2:
                try:
                    quad = item[1]
                    points = (
                        quad.ul,
                        quad.ur,
                        quad.lr,
                        quad.ll,
                    )
                    for first, second in zip(
                        points,
                        points[1:] + points[:1],
                    ):
                        add_segment(
                            float(first.x),
                            float(first.y),
                            float(second.x),
                            float(second.y),
                        )
                except (AttributeError, TypeError, ValueError):
                    continue
    return tuple(horizontal), tuple(vertical)


def _cluster_coordinates(values, tolerance=2.0):
    clusters = []
    for value in sorted(values):
        if clusters and abs(value - clusters[-1][-1]) <= tolerance:
            clusters[-1].append(value)
        else:
            clusters.append([value])
    return tuple(sum(cluster) / len(cluster) for cluster in clusters)


def _has_grid_like_drawings(drawings, excluded_rects=()) -> bool:
    horizontal, vertical = _drawing_axis_segments(
        drawings,
        excluded_rects=excluded_rects,
    )
    horizontal_levels = _cluster_coordinates(segment[2] for segment in horizontal)
    vertical_levels = _cluster_coordinates(segment[0] for segment in vertical)
    if not (
        (len(horizontal_levels) >= 3 and len(vertical_levels) >= 2)
        or (len(horizontal_levels) >= 2 and len(vertical_levels) >= 3)
    ):
        return False

    intersections = set()
    for vertical_x, vertical_y0, vertical_y1 in vertical:
        for horizontal_x0, horizontal_x1, horizontal_y in horizontal:
            if (
                horizontal_x0 - 2 <= vertical_x <= horizontal_x1 + 2
                and vertical_y0 - 2 <= horizontal_y <= vertical_y1 + 2
            ):
                intersections.add((round(vertical_x, 1), round(horizontal_y, 1)))
    return len(intersections) >= 6


def _page_has_table_like_layout(
    page,
    drawings,
    excluded_rects=(),
    allow_single_row=True,
) -> bool:
    return _has_grid_like_drawings(
        drawings,
        excluded_rects=excluded_rects,
    ) or _has_repeated_text_columns(
        page,
        excluded_rects=excluded_rects,
        allow_single_row=allow_single_row,
    )


def analyze_pdf_fidelity_risk(
    path: str | PathLike[str],
    max_pages: int | None = 3,
    vector_threshold: int = 120,
    cancel_requested=None,
    detailed_table_page_limit: int | None = None,
) -> PdfFidelityRisk:
    """Inspect the requested PDF pages for conversion-sensitive structures.

    Text, font-mapping, form, vector, and layout risks are scanned on every
    requested page. Documents within ``detailed_table_page_limit`` receive full
    table extraction and cell geometry analysis. Longer documents use the fast
    all-page scan and report ``table_analysis_limited`` so callers can continue
    with an editable engine and present an explicit completion warning.
    """

    if (
        max_pages is not None
        and (
            isinstance(max_pages, bool)
            or not isinstance(max_pages, int)
            or max_pages <= 0
        )
    ):
        raise ValueError("max_pages 必须是大于 0 的整数")
    if (
        isinstance(vector_threshold, bool)
        or not isinstance(vector_threshold, int)
        or vector_threshold < 0
    ):
        raise ValueError("vector_threshold 必须是大于或等于 0 的整数")

    if cancel_requested is not None and not callable(cancel_requested):
        raise ValueError("cancel_requested must be callable")
    if (
        detailed_table_page_limit is not None
        and (
            isinstance(detailed_table_page_limit, bool)
            or not isinstance(detailed_table_page_limit, int)
            or detailed_table_page_limit <= 0
        )
    ):
        raise ValueError("detailed_table_page_limit 必须是大于 0 的整数")
    _raise_if_pdf_analysis_cancelled(cancel_requested)

    source_path = Path(path)
    document = None
    try:
        document = pymupdf.open(source_path)
        _raise_if_pdf_analysis_cancelled(cancel_requested)
        page_count = len(document)
        analyzed_pages = page_count if max_pages is None else min(page_count, max_pages)
        effective_detail_limit = (
            _DETAILED_TABLE_ANALYSIS_PAGE_LIMIT
            if detailed_table_page_limit is None
            else detailed_table_page_limit
        )
        table_analysis_limited = analyzed_pages > effective_detail_limit
        table_count = 0
        table_text_character_count = 0
        table_texts = []
        table_cell_matrices = []
        table_full_width_span_rows = []
        table_shapes = []
        table_cell_counts = []
        table_cell_spans = []
        table_column_widths = []
        table_cell_border_edges = []
        table_text_extraction_failure_count = 0
        widget_count = 0
        checkbox_symbol_count = 0
        vector_mark_count = 0
        symbol_font_run_count = 0
        selectable_text_character_count = 0
        selectable_text_parts = []
        unverifiable_text_character_count = 0
        vector_path_count = 0
        _raise_if_pdf_analysis_cancelled(cancel_requested)
        tagged_table_structure_present = _document_has_tagged_table_structure(
            document
        )
        _raise_if_pdf_analysis_cancelled(cancel_requested)
        table_layout_suspected = tagged_table_structure_present
        unconfirmed_table_region_count = 0

        for page_number in range(analyzed_pages):
            _raise_if_pdf_analysis_cancelled(cancel_requested)
            page = document[page_number]
            (
                page_checkbox_count,
                page_symbol_font_runs,
                page_selectable_text_count,
                page_selectable_text,
                page_mapping_warning_count,
            ) = _page_text_risks(page)
            checkbox_symbol_count += page_checkbox_count
            symbol_font_run_count += page_symbol_font_runs
            selectable_text_character_count += page_selectable_text_count
            selectable_text_parts.append(page_selectable_text)
            unverifiable_text_character_count += page_mapping_warning_count
            _raise_if_pdf_analysis_cancelled(cancel_requested)
            page_drawings = tuple(page.get_drawings() or ())
            _raise_if_pdf_analysis_cancelled(cancel_requested)
            vector_mark_count += _small_vector_mark_count(page_drawings)
            page_layout_suspected = _page_has_table_like_layout(
                page,
                page_drawings,
            )
            if page_layout_suspected:
                table_layout_suspected = True

            find_tables = getattr(page, "find_tables", None)
            if callable(find_tables) and not table_analysis_limited:
                _raise_if_pdf_analysis_cancelled(cancel_requested)
                default_finder = find_tables()
                _raise_if_pdf_analysis_cancelled(cancel_requested)
                default_tables = tuple(
                    getattr(default_finder, "tables", ()) or ()
                )
                table_extracted_rows = {}
                for table in default_tables:
                    try:
                        table_extracted_rows[id(table)] = tuple(
                            table.extract() or ()
                        )
                    except Exception:
                        table_extracted_rows[id(table)] = None

                _raise_if_pdf_analysis_cancelled(cancel_requested)
                try:
                    precise_finder = find_tables(snap_x_tolerance=1.0)
                    _raise_if_pdf_analysis_cancelled(cancel_requested)
                except (TypeError, ValueError):
                    precise_tables = ()
                else:
                    precise_tables = tuple(
                        getattr(precise_finder, "tables", ()) or ()
                    )
                    for table in precise_tables:
                        if id(table) in table_extracted_rows:
                            continue
                        try:
                            table_extracted_rows[id(table)] = tuple(
                                table.extract() or ()
                            )
                        except Exception:
                            table_extracted_rows[id(table)] = None

                tables = _select_precise_line_tables(
                    default_tables,
                    precise_tables,
                    table_extracted_rows,
                )
                _raise_if_pdf_analysis_cancelled(cancel_requested)
                text_table_ids = set()
                if page_selectable_text_count:
                    clips = _table_exclusion_clips(page, tables)
                    text_tables = _find_text_table_candidates(
                        page,
                        find_tables,
                        clips,
                        extracted_rows_by_id=table_extracted_rows,
                        cancel_requested=cancel_requested,
                    )
                    _raise_if_pdf_analysis_cancelled(cancel_requested)
                    text_table_ids = {id(table) for table in text_tables}
                    tables = _deduplicate_table_candidates(
                        tables + text_tables,
                        extracted_rows_by_id=table_extracted_rows,
                    )
                table_count += len(tables)
                for table in tables:
                    _raise_if_pdf_analysis_cancelled(cancel_requested)
                    bbox_text = _table_bbox_text(page, table)
                    bbox_character_count = sum(
                        not character.isspace()
                        for character in bbox_text
                    )
                    table_uncertain = getattr(table, "bbox", None) is None
                    extracted_rows = ()
                    try:
                        if id(table) in table_extracted_rows:
                            extracted_rows = table_extracted_rows[id(table)]
                            if extracted_rows is None:
                                raise RuntimeError("表格单元格提取失败")
                        else:
                            extracted_rows = tuple(table.extract() or ())
                    except Exception:
                        extracted_rows = ()
                        table_uncertain = True
                        table_text = bbox_text
                        baseline_character_count = bbox_character_count
                        table_shape = _effective_table_shape(
                            table,
                            (),
                            normalize_text_strategy_rows=id(table) in text_table_ids,
                        )
                    else:
                        extracted_character_count = 0
                        extracted_text_parts = []
                        for row in extracted_rows:
                            for cell_text in row or ():
                                text_value = cell_text or ""
                                extracted_text_parts.append(text_value)
                                extracted_character_count += sum(
                                    not character.isspace()
                                    for character in text_value
                                )

                        table_text = "".join(extracted_text_parts)
                        baseline_character_count = extracted_character_count
                        table_uncertain = (
                            not extracted_character_count
                            or bbox_character_count > extracted_character_count
                        )
                        if bbox_character_count > extracted_character_count:
                            table_text = bbox_text
                            baseline_character_count = bbox_character_count
                        table_shape = _effective_table_shape(
                            table,
                            extracted_rows,
                            normalize_text_strategy_rows=id(table) in text_table_ids,
                        )

                    normalize_text_rows = id(table) in text_table_ids
                    indexed_matrix_rows = tuple(
                        (
                            row_index,
                            tuple(
                                None if cell_text is None else str(cell_text)
                                for cell_text in (row or ())
                            ),
                        )
                        for row_index, row in enumerate(extracted_rows)
                    )
                    if normalize_text_rows:
                        indexed_matrix_rows = tuple(
                            (row_index, row)
                            for row_index, row in indexed_matrix_rows
                            if any(
                                str(cell_text or "").strip()
                                for cell_text in row
                            )
                        )
                    retained_row_indices = tuple(
                        row_index for row_index, _row in indexed_matrix_rows
                    )
                    matrix_rows = tuple(
                        row for _row_index, row in indexed_matrix_rows
                    )
                    full_width_span_rows = _table_full_width_span_flags(
                        table,
                        retained_row_indices,
                        table_shape[0],
                    )
                    table_cell_count = _effective_table_cell_count(
                        table,
                        extracted_rows,
                        normalize_text_strategy_rows=normalize_text_rows,
                    )
                    geometry = _table_geometry_model(table, table_shape)
                    if geometry is None:
                        cell_spans = _cell_spans_from_matrix(
                            matrix_rows,
                            table_shape,
                        )
                        column_widths = ()
                    else:
                        cell_spans, column_widths = geometry
                    cell_spans = tuple(sorted(cell_spans or ()))
                    border_edges = _table_cell_border_edges_from_drawings(
                        page,
                        table,
                        cell_spans,
                        page_drawings,
                    )
                    if (
                        table_shape[0] < 1
                        or table_shape[1] < 1
                        or table_cell_count < 1
                        or len(cell_spans) != table_cell_count
                        or len(border_edges) != len(cell_spans)
                        or len(column_widths) != table_shape[1]
                    ):
                        table_uncertain = True
                    if table_uncertain:
                        table_text_extraction_failure_count += 1
                    table_texts.append(table_text)
                    table_cell_matrices.append(matrix_rows)
                    table_full_width_span_rows.append(full_width_span_rows)
                    table_shapes.append(table_shape)
                    table_cell_counts.append(table_cell_count)
                    table_cell_spans.append(cell_spans)
                    table_column_widths.append(tuple(column_widths or ()))
                    table_cell_border_edges.append(border_edges)
                    table_text_character_count += baseline_character_count

                table_rects = _table_layout_exclusion_rects(page, tables)
                if page_layout_suspected and (
                    not tables
                    or _page_has_table_like_layout(
                        page,
                        page_drawings,
                        excluded_rects=table_rects,
                        allow_single_row=False,
                    )
                ):
                    unconfirmed_table_region_count += 1

            elif page_layout_suspected:
                unconfirmed_table_region_count += 1

            _raise_if_pdf_analysis_cancelled(cancel_requested)
            widgets = page.widgets()
            widget_count += sum(1 for _widget in (widgets or ()))
            vector_path_count += len(page_drawings)
            _raise_if_pdf_analysis_cancelled(cancel_requested)

        reasons: list[str] = []
        if table_count:
            reasons.append(f"检测到表格结构（{table_count} 个）")
        if tagged_table_structure_present:
            reasons.append("PDF 标签结构声明存在表格")
        if tagged_table_structure_present and not table_count:
            unconfirmed_table_region_count = max(
                1,
                unconfirmed_table_region_count,
            )
        if unconfirmed_table_region_count:
            reasons.append(
                "版面存在无法确认的表格区域，表格结构无法确认"
                f"（{unconfirmed_table_region_count} 处）"
            )
        if table_analysis_limited:
            reasons.append(
                f"长文档已完成全部 {analyzed_pages} 页快速文字检查，"
                "未执行逐页表格结构重建"
            )
        if widget_count:
            reasons.append(f"检测到 PDF 表单控件（{widget_count} 个）")
        if table_text_extraction_failure_count:
            reasons.append(
                f"有 {table_text_extraction_failure_count} 个表格的单元格提取不完整"
            )
        if checkbox_symbol_count:
            reasons.append(f"检测到显式复选或勾选符号（{checkbox_symbol_count} 个）")
        if vector_mark_count:
            reasons.append(
                f"检测到疑似矢量复选或勾选标记（{vector_mark_count} 处）"
            )
        if symbol_font_run_count:
            reasons.append(
                "检测到 Wingdings/Webdings/Symbol 字体文本"
                f"（{symbol_font_run_count} 处）"
            )
        if unverifiable_text_character_count:
            reasons.append(
                "检测到无法可靠映射的文字字符"
                f"（{unverifiable_text_character_count} 个）"
            )
        if vector_path_count > vector_threshold:
            reasons.append(
                f"矢量路径较多（{vector_path_count} 条，阈值 {vector_threshold}）"
            )

        return PdfFidelityRisk(
            path=source_path,
            page_count=page_count,
            analyzed_pages=analyzed_pages,
            table_count=table_count,
            table_text_character_count=table_text_character_count,
            table_texts=tuple(table_texts),
            table_cell_matrices=tuple(table_cell_matrices),
            table_full_width_span_rows=tuple(table_full_width_span_rows),
            table_shapes=tuple(table_shapes),
            table_cell_counts=tuple(table_cell_counts),
            table_cell_spans=tuple(table_cell_spans),
            table_column_widths=tuple(table_column_widths),
            table_cell_border_edges=tuple(table_cell_border_edges),
            # A confirmed empty grid is still an editable table and must never
            # be allowed to fall through to full-page image conversion.
            editable_table_candidate=bool(table_count),
            table_layout_suspected=table_layout_suspected,
            unconfirmed_table_region_count=unconfirmed_table_region_count,
            tagged_table_structure_present=tagged_table_structure_present,
            widget_count=widget_count,
            checkbox_symbol_count=checkbox_symbol_count,
            vector_mark_count=vector_mark_count,
            table_text_extraction_failure_count=table_text_extraction_failure_count,
            selectable_text_character_count=selectable_text_character_count,
            selectable_text="".join(selectable_text_parts),
            unverifiable_text_character_count=unverifiable_text_character_count,
            table_analysis_limited=table_analysis_limited,
            symbol_font_run_count=symbol_font_run_count,
            vector_path_count=vector_path_count,
            is_complex=bool(reasons),
            reasons=tuple(reasons),
        )
    except PdfFidelityAnalysisCancelled:
        raise
    except Exception as exc:
        raise RuntimeError(f"无法分析 PDF“{source_path}”：{exc}") from exc
    finally:
        if document is not None:
            document.close()
