"""The workbook JSON contract shared by the LLM, the API, the frontend grid and the builder.

The LLM is constrained to emit exactly `WorkbookSpec` (via structured outputs), the
frontend renders and edits the same structure, and `workbook_builder` compiles it
into a real .xlsx. Keeping one schema means the preview and the download never drift.

Conventions
-----------
* Row 1 of every sheet is the header row built from `columns[*].header`.
  Data rows start at Excel row 2, so `rows[0]` is row 2, `rows[1]` is row 3, etc.
* A cell whose value is a string starting with "=" is an Excel formula,
  written with A1 references (e.g. "=SUM(B2:B10)", "='Income'!C5*0.15").
* `columns[i].format` drives number formatting for the whole column; a cell can
  override it with `CellSpec.format`.
"""

from __future__ import annotations

from enum import Enum

from pydantic import BaseModel, Field


class NumberFormat(str, Enum):
    text = "text"
    integer = "integer"
    decimal = "decimal"
    currency = "currency"
    percent = "percent"
    date = "date"


# Supported currency codes and their symbols (the frontend uses the same table). Unknown codes fall back to USD.
CURRENCY_SYMBOLS = {
    "USD": "$", "EUR": "€", "GBP": "£", "JPY": "¥", "INR": "₹", "CNY": "CN¥",
    "CAD": "CA$", "AUD": "A$", "CHF": "CHF ", "MXN": "MX$", "BRL": "R$", "ZAR": "R",
}


class RowKind(str, Enum):
    data = "data"
    subtotal = "subtotal"
    total = "total"
    section = "section"  # a label row that groups the rows below it


class CondOperator(str, Enum):
    greater_than = "greater_than"
    less_than = "less_than"
    between = "between"
    equal = "equal"
    negative = "negative"
    contains = "contains"  # text contains `text`, case-insensitive
    color_scale = "color_scale"  # 3-colour heatmap, low red -> high green (`reverse` flips it)
    data_bar = "data_bar"  # in-cell bars sized by value


class CondStyle(str, Enum):
    good = "good"  # green fill
    bad = "bad"  # red fill
    warning = "warning"  # amber fill
    custom = "custom"  # use the rule's fill_color / font_color


class Align(str, Enum):
    left = "left"
    center = "center"
    right = "right"


class BorderStyle(str, Enum):
    none = "none"
    thin = "thin"
    thick = "thick"


class CellSpec(BaseModel):
    """One cell. `value` is a literal or, when it starts with '=', a formula.

    Style fields are optional overrides on top of the automatic theme; leave them
    empty to get the default corporate styling.
    """

    value: str | float | bool | None = None
    format: NumberFormat | None = Field(
        default=None, description="Overrides the column's number format for this cell only."
    )
    bold: bool = False
    italic: bool = False
    underline: bool = False
    font_color: str | None = Field(default=None, description='Hex colour such as "#C00000".')
    fill_color: str | None = Field(default=None, description='Hex background colour such as "#FFF2CC".')
    align: Align | None = None
    border: BorderStyle | None = Field(default=None, description="Overrides the default thin grid border.")
    note: str | None = Field(default=None, description="Optional cell comment shown in Excel.")
    decimals: int | None = Field(
        default=None, description="Decimal places (0-10) for decimal/currency/percent formats. Overrides the column."
    )
    currency: str | None = Field(default=None, description='ISO code such as "EUR" when the format is currency. Overrides the column.')
    wrap: bool = Field(default=False, description="Wrap long text onto several lines inside the cell.")
    ai_value: str | None = Field(
        default=None, description="Last result of an =AI(...) formula in this cell; downloads write it as the value."
    )


class RowSpec(BaseModel):
    kind: RowKind = RowKind.data
    cells: list[CellSpec] = Field(
        description="Cells left to right, aligned with the sheet's columns. May be shorter than columns."
    )
    height: float | None = Field(default=None, description="Row height in points. Omit for the default.")


class ColumnSpec(BaseModel):
    header: str
    format: NumberFormat = NumberFormat.text
    width: float | None = Field(default=None, description="Character width. Omit to auto-fit.")
    decimals: int | None = Field(
        default=None, description="Decimal places (0-10) for decimal/currency/percent formats. Omit for the default."
    )
    currency: str | None = Field(default=None, description='ISO code such as "EUR" when the format is currency. Omit for USD.')


class ConditionalFormatSpec(BaseModel):
    range: str = Field(description='A1 range on this sheet, e.g. "D2:D13".')
    operator: CondOperator
    value1: float | None = None
    value2: float | None = Field(default=None, description="Upper bound, only for 'between'.")
    text: str | None = Field(default=None, description="Text to look for, only for 'contains'.")
    reverse: bool = Field(default=False, description="color_scale only: low values green, high values red.")
    style: CondStyle = Field(default=CondStyle.good, description="Ignored for color_scale and data_bar.")
    fill_color: str | None = Field(default=None, description='Hex fill for style "custom" or the data_bar colour, e.g. "#FFC7CE".')
    font_color: str | None = Field(default=None, description='Hex font colour for style "custom".')


class ChartType(str, Enum):
    column = "column"  # vertical bars
    bar = "bar"  # horizontal bars
    line = "line"
    pie = "pie"
    scatter = "scatter"  # X/Y points; `categories` is the X values range
    area = "area"
    doughnut = "doughnut"


class SeriesKind(str, Enum):
    column = "column"
    line = "line"


class ChartSeriesSpec(BaseModel):
    name: str
    range: str = Field(description='Values on this sheet, one column or row, e.g. "D2:D13".')
    color: str | None = Field(default=None, description='Hex colour such as "#4472C4". Omit for the theme colour.')
    kind: SeriesKind | None = Field(
        default=None, description='Column charts only: "line" draws this series as a line on the same axes (combo chart).'
    )


class ChartSpec(BaseModel):
    type: ChartType
    title: str
    categories: str = Field(description='Category labels on this sheet, e.g. "A2:A13". For scatter, the X values.')
    series: list[ChartSeriesSpec] = Field(description="One or more value series. Pie and doughnut charts use only the first.")
    anchor: str | None = Field(default=None, description='Top-left cell for the chart, e.g. "H2". Omit to place it right of the data.')
    stacked: bool = Field(default=False, description="Stack the series (column, bar and area charts).")
    x_title: str | None = Field(default=None, description="Horizontal axis title.")
    y_title: str | None = Field(default=None, description="Vertical axis title.")


class ValidationType(str, Enum):
    list = "list"  # dropdown
    whole = "whole"  # whole number between min and max
    decimal = "decimal"
    date = "date"


class DataValidationSpec(BaseModel):
    range: str = Field(description='A1 range on this sheet, e.g. "E2:E50".')
    type: ValidationType
    options: list[str] = Field(default_factory=list, description="Dropdown choices, only for type 'list'.")
    min: float | None = None
    max: float | None = None
    message: str | None = Field(default=None, description="Shown when someone enters an invalid value.")


class FilterSpec(BaseModel):
    """Show only data rows whose displayed text in `column` is one of `values`."""

    column: int = Field(description="0-based column index.")
    values: list[str]


class SheetSpec(BaseModel):
    name: str = Field(description="Tab name, at most 31 characters, none of []:*?/\\")
    columns: list[ColumnSpec]
    rows: list[RowSpec]
    freeze_header: bool = True
    freeze_columns: int = Field(default=0, description="Leading columns kept visible when scrolling sideways.")
    show_gridlines: bool = True
    conditional_formats: list[ConditionalFormatSpec] = Field(default_factory=list)
    charts: list[ChartSpec] = Field(default_factory=list)
    validations: list[DataValidationSpec] = Field(default_factory=list)
    merges: list[str] = Field(
        default_factory=list,
        description='A1 ranges to merge, e.g. "A1:C1". The value and style come from the top-left cell.',
    )
    filters: list[FilterSpec] = Field(default_factory=list, description="Column filters that hide non-matching data rows.")
    header_height: float | None = Field(default=None, description="Header row height in points. Omit for the default.")
    locked: bool = Field(default=False, description="Protect the whole sheet (read-only).")
    locked_ranges: list[str] = Field(
        default_factory=list, description='Protect only these A1 ranges, e.g. "A1:H1"; everything else stays editable.'
    )


class ScenarioSpec(BaseModel):
    """A named set of input values, keyed by absolute cell address such as "'Inputs'!B2"."""

    name: str
    values: dict[str, float] = Field(default_factory=dict)


class TileKind(str, Enum):
    kpi = "kpi"
    chart = "chart"
    text = "text"


class DashboardTile(BaseModel):
    """A tile on the dashboard page. Positions are grid units on a 12-column layout."""

    id: str
    kind: TileKind
    title: str = ""
    ref: str | None = Field(default=None, description='kpi: the cell shown, e.g. "Summary!B4".')
    compare_ref: str | None = Field(default=None, description='kpi: optional comparison cell, e.g. "Summary!C4".')
    sheet: str | None = Field(default=None, description="chart: the sheet holding the chart.")
    chart: int | None = Field(default=None, description="chart: index in that sheet's charts.")
    text: str | None = Field(default=None, description="text: the note shown.")
    x: int = 0
    y: int = 0
    w: int = 3
    h: int = 2


class DashboardSpec(BaseModel):
    title: str = "Dashboard"
    tiles: list[DashboardTile] = Field(default_factory=list)


class ConnectionKind(str, Enum):
    csv = "csv"
    gsheet = "gsheet"
    json = "json"


class DataConnection(BaseModel):
    """Data pulled from a URL into a sheet, refreshed by hand or every `refresh_minutes`."""

    id: str
    name: str = ""
    kind: ConnectionKind
    url: str
    path: str | None = Field(default=None, description='JSON only: dotted path to the list of records, e.g. "data.items".')
    sheet: str = Field(description="Sheet the data is written into (created if missing); its rows are replaced.")
    refresh_minutes: int | None = None
    last_refreshed: int | None = Field(default=None, description="ms since epoch.")


class WorkbookSpec(BaseModel):
    title: str
    sheets: list[SheetSpec]
    scenarios: list[ScenarioSpec] = Field(default_factory=list)
    dashboard: DashboardSpec | None = None
    connections: list[DataConnection] = Field(default_factory=list)


# ---------------------------------------------------------------------------
# API request / response models
# ---------------------------------------------------------------------------


class Issue(BaseModel):
    """A validation finding. `fixed` means the engine auto-corrected it."""

    severity: str  # "error" | "warning"
    sheet: str | None = None
    cell: str | None = None
    message: str
    fixed: bool = False


class GenerateRequest(BaseModel):
    prompt: str = Field(min_length=1)


class EditRequest(BaseModel):
    prompt: str = Field(min_length=1)
    workbook: WorkbookSpec


class WorkbookResponse(BaseModel):
    workbook: WorkbookSpec
    summary: str = Field(description="One or two sentences on what changed, shown in the chat.")
    issues: list[Issue] = Field(default_factory=list)


class ExplainRequest(BaseModel):
    workbook: WorkbookSpec
    sheet: str
    cell: str = Field(description='A1 address, e.g. "E4".')
    question: str | None = None
    computed: str | None = Field(default=None, description="What the cell currently shows, e.g. '#DIV/0!'.")


class ExplainResult(BaseModel):
    """What the model returns for a single-cell question."""

    explanation: str = Field(description="Plain-English answer, a few sentences at most.")
    fixed_value: str | None = Field(
        default=None, description="A corrected formula or value for the cell, only when it is wrong."
    )


class Aggregation(str, Enum):
    sum = "sum"
    average = "average"
    count = "count"
    min = "min"
    max = "max"


class PivotRequest(BaseModel):
    workbook: WorkbookSpec
    source_sheet: str
    row_field: str = Field(description="Header of the column to group rows by.")
    column_field: str | None = Field(default=None, description="Optional header to spread across columns.")
    value_field: str | None = Field(default=None, description="Header of the numbers to aggregate. Not needed for count.")
    aggregation: Aggregation = Aggregation.sum
    name: str | None = Field(default=None, description="Name for the new summary sheet.")


class AskRequest(BaseModel):
    workbook: WorkbookSpec
    question: str = Field(min_length=1)
    values: dict[str, list[list[str]]] = Field(
        default_factory=dict,
        description="Per sheet, what the grid displays (formulas calculated). Row 0 is the header row (Excel row 1).",
    )


class CellReference(BaseModel):
    sheet: str
    range: str = Field(description='A1 cell or range on that sheet, e.g. "D7" or "B2:B13".')


class AskResult(BaseModel):
    """What the model returns for a question about the data."""

    answer: str = Field(description="A plain answer in one to four sentences, using the actual numbers.")
    references: list[CellReference] = Field(
        default_factory=list, description="Up to 5 cells or ranges that back up the answer."
    )


class CheckRequest(BaseModel):
    workbook: WorkbookSpec
    values: dict[str, list[list[str]]] = Field(
        default_factory=dict,
        description="Per sheet, what the grid displays (formulas calculated). Row 0 is the header row (Excel row 1).",
    )
    use_ai: bool = True


class Finding(BaseModel):
    id: str = Field(description="Stable id derived from kind, sheet and cell.")
    severity: str  # "error" | "warning" | "info"
    sheet: str
    cell: str = Field(description='A1 address, e.g. "D7".')
    kind: str
    message: str
    fix: str | None = Field(default=None, description='Replacement value or formula (starting with "=").')


class CheckResponse(BaseModel):
    summary: str
    findings: list[Finding] = Field(default_factory=list)


class AIFinding(BaseModel):
    """A problem the model spotted while reviewing the workbook."""

    severity: str = Field(description='"error", "warning" or "info".')
    sheet: str
    cell: str = Field(description='A1 address of the cell to change, e.g. "D7".')
    kind: str = Field(description='Short label such as "range_misses_rows", "double_count" or "wrong_percentage".')
    message: str = Field(description="One or two plain-English sentences: what is wrong and why it matters.")
    fix: str | None = Field(default=None, description='Corrected formula (starting with "=") or value for that cell.')


class AICheckResult(BaseModel):
    """What the model returns for a workbook review."""

    summary: str = Field(description="One sentence on the overall health of the workbook.")
    findings: list[AIFinding] = Field(default_factory=list, description="Up to 8 problems, most serious first.")


class LLMWorkbookResult(BaseModel):
    """What the model returns: the full workbook plus a chat-facing summary."""

    summary: str
    workbook: WorkbookSpec


# ---------------------------------------------------------------------------
# AI fill, =AI() evaluation and photo / PDF extraction
# ---------------------------------------------------------------------------


class FillOutput(str, Enum):
    text = "text"
    number = "number"


class AIFillRequest(BaseModel):
    instruction: str = Field(min_length=1)
    headers: list[str] = Field(default_factory=list)
    rows: list[list[str | float | bool | None]] = Field(default_factory=list)
    output: FillOutput = FillOutput.text


class AIFillResponse(BaseModel):
    values: list[str | float | None]


class AIFillValue(BaseModel):
    row: int = Field(description="0-based index of the input row.")
    value: str = Field(description="The cell value: a short text, or a plain number for number output. Empty if unknown.")


class AIFillResult(BaseModel):
    """What the model returns for a fill request."""

    values: list[AIFillValue] = Field(description="One entry per input row.")


class AIEvalItem(BaseModel):
    key: str = Field(min_length=1, max_length=200)
    prompt: str
    inputs: list[str | float | bool | None] = Field(default_factory=list)


class AIEvalRequest(BaseModel):
    items: list[AIEvalItem]


class AIEvalResponse(BaseModel):
    results: dict[str, str]


class AIEvalAnswer(BaseModel):
    key: str
    value: str = Field(description="The cell value: short, no explanation.")


class AIEvalResult(BaseModel):
    """What the model returns for a batch of =AI() cells."""

    results: list[AIEvalAnswer]


class ExtractedColumn(BaseModel):
    header: str
    format: NumberFormat = Field(description="currency for money, percent for rates, date for dates, integer/decimal for other numbers, else text.")
    currency: str | None = Field(default=None, description='ISO code such as "EUR" for currency columns. Omit for USD.')


class ExtractedRow(BaseModel):
    kind: RowKind = RowKind.data
    cells: list[str | float | None] = Field(
        description="Left to right, one per column. Numbers as JSON numbers (percent as a fraction), dates as YYYY-MM-DD, blanks as null."
    )


class ExtractedTable(BaseModel):
    name: str = Field(description="Short tab name, at most 31 characters.")
    columns: list[ExtractedColumn]
    rows: list[ExtractedRow]


class ExtractResult(BaseModel):
    """What the model returns for a photo or PDF of tables."""

    summary: str = Field(description="One sentence on what was found, mentioning anything that was unreadable.")
    title: str = Field(description="A short workbook title.")
    tables: list[ExtractedTable] = Field(description="Every table found, at most 5.")


class ReportSummary(BaseModel):
    summary: str = Field(description="Two or three plain sentences on what the numbers show.")
