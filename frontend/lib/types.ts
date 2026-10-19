// Mirrors backend/app/schemas.py. Keep the two in sync.

export type NumberFormat = "text" | "integer" | "decimal" | "currency" | "percent" | "date";
export type RowKind = "data" | "subtotal" | "total" | "section";
export type Align = "left" | "center" | "right";
export type BorderStyle = "none" | "thin" | "thick";
export type CellValue = string | number | boolean | null;

export interface CellSpec {
  value?: CellValue;
  format?: NumberFormat | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  font_color?: string | null;
  fill_color?: string | null;
  align?: Align | null;
  border?: BorderStyle | null;
  note?: string | null;
  /** Decimal places for decimal / currency / percent formats. */
  decimals?: number | null;
  /** ISO currency code for the currency format (USD when empty). */
  currency?: string | null;
  wrap?: boolean;
  /** The last result of an =AI(...) formula in this cell; downloads write it as the value. */
  ai_value?: string | null;
}

export interface RowSpec {
  kind?: RowKind;
  /** Row height in points. */
  height?: number | null;
  cells: CellSpec[];
}

export interface ColumnSpec {
  header: string;
  format?: NumberFormat;
  width?: number | null;
  decimals?: number | null;
  currency?: string | null;
}

export type CondOperator = "greater_than" | "less_than" | "between" | "equal" | "negative" | "contains" | "color_scale" | "data_bar";
export type CondStyle = "good" | "bad" | "warning" | "custom";

export interface ConditionalFormatSpec {
  range: string;
  operator: CondOperator;
  value1?: number | null;
  value2?: number | null;
  /** Text to look for, for "contains". */
  text?: string | null;
  /** Colour scale: low values green instead of red. */
  reverse?: boolean;
  style?: CondStyle;
  fill_color?: string | null;
  font_color?: string | null;
}

export type ChartType = "column" | "bar" | "line" | "pie" | "scatter" | "area" | "doughnut";

export interface ChartSeriesSpec {
  name: string;
  range: string;
  color?: string | null;
  /** In a column chart, "line" draws this series as a line (a combo chart). */
  kind?: "column" | "line" | null;
}

export interface ChartSpec {
  type: ChartType;
  title: string;
  categories: string;
  series: ChartSeriesSpec[];
  anchor?: string | null;
  stacked?: boolean;
  x_title?: string | null;
  y_title?: string | null;
}

export type ValidationType = "list" | "whole" | "decimal" | "date";

export interface DataValidationSpec {
  range: string;
  type: ValidationType;
  options?: string[];
  min?: number | null;
  max?: number | null;
  message?: string | null;
}

export interface SheetSpec {
  name: string;
  columns: ColumnSpec[];
  rows: RowSpec[];
  freeze_header?: boolean;
  /** Leading columns kept visible when scrolling sideways. */
  freeze_columns?: number;
  show_gridlines?: boolean;
  conditional_formats?: ConditionalFormatSpec[];
  charts?: ChartSpec[];
  validations?: DataValidationSpec[];
  /** Ranges shown as one cell, e.g. "A1:C1". Value and style come from the top-left cell. */
  merges?: string[];
  /** Show only data rows whose displayed text in `column` is one of `values`. */
  filters?: FilterSpec[];
  header_height?: number | null;
  /** Protection: the whole sheet is read-only for editors and viewers. */
  locked?: boolean;
  /** Protection: these ranges (e.g. "A1:H1") are read-only for everyone but the owner. */
  locked_ranges?: string[];
}

export interface FilterSpec {
  column: number;
  values: string[];
}

export interface ScenarioSpec {
  name: string;
  /** Input values keyed by address, e.g. "Inputs!B2". */
  values: Record<string, number>;
}

export interface WorkbookSpec {
  title: string;
  sheets: SheetSpec[];
  scenarios?: ScenarioSpec[];
  dashboard?: DashboardSpec | null;
  connections?: DataConnection[];
}

/** A tile on the dashboard page. Positions are in grid units on a 12-column layout. */
export interface DashboardTile {
  id: string;
  kind: "kpi" | "chart" | "text";
  title: string;
  /** kpi: the cell whose value is shown, e.g. "Summary!B4". */
  ref?: string | null;
  /** kpi: optional comparison cell, shown as a change, e.g. "Summary!C4". */
  compare_ref?: string | null;
  /** chart: which chart, by sheet name and index in that sheet's charts. */
  sheet?: string | null;
  chart?: number | null;
  /** text: the note shown. */
  text?: string | null;
  x: number;
  y: number;
  w: number;
  h: number;
}

export interface DashboardSpec {
  title: string;
  tiles: DashboardTile[];
}

/** Data pulled from outside into a sheet, refreshed on demand or on a schedule. */
export interface DataConnection {
  id: string;
  name: string;
  kind: "csv" | "gsheet" | "json";
  url: string;
  /** For JSON: dotted path to the array of records, e.g. "data.items". */
  path?: string | null;
  /** The sheet the data is written into (created if missing). Its rows are replaced on refresh. */
  sheet: string;
  /** Minutes between automatic refreshes; null means only by hand. */
  refresh_minutes?: number | null;
  last_refreshed?: number | null;
}

export interface Finding {
  id: string;
  severity: "error" | "warning" | "info";
  sheet: string;
  cell: string;
  kind: string;
  message: string;
  fix?: string | null;
}

export interface Issue {
  severity: "error" | "warning";
  sheet?: string | null;
  cell?: string | null;
  message: string;
  fixed: boolean;
}

export interface WorkbookResponse {
  workbook: WorkbookSpec;
  summary: string;
  issues: Issue[];
}

export interface ChatMessage {
  id: string;
  role: "user" | "assistant";
  text: string;
  issues?: Issue[];
  error?: boolean;
  /** The version this message produced, so it can be restored from the chat. */
  checkpointId?: string;
  /** Cells an answer refers to, shown as links that select them. */
  references?: { sheet: string; range: string }[];
  /** "question" / "answer" for Ask mode, so they read differently from changes. */
  kind?: "question" | "answer";
}

export interface Checkpoint {
  id: string;
  label: string;
  at: number;
  workbook: WorkbookSpec;
}

export interface SavedDoc {
  id: string;
  title: string;
  updated: number;
  workbook: WorkbookSpec;
  messages: ChatMessage[];
  checkpoints: Checkpoint[];
}

export interface TemplateInfo {
  id: string;
  title: string;
  description: string;
  category: string;
  sheets: number;
}

export interface AskResult {
  answer: string;
  references: { sheet: string; range: string }[];
}

export interface ExplainResult {
  explanation: string;
  fixed_value?: string | null;
}

export type Aggregation = "sum" | "average" | "count" | "min" | "max";

export interface PivotOptions {
  source_sheet: string;
  row_field: string;
  column_field?: string | null;
  value_field?: string | null;
  aggregation: Aggregation;
}

/** A rectangular selection in grid coordinates (row 0 is the header row). */
export interface Selection {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}
