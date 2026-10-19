// Checks for the clean-up and what-if tools:  npx tsx scripts/check-tools.ts
import { applyCleanup, changeCaseText, cleanupScope, duplicateRows, parseDateText, parseNumberText, trimText } from "../lib/cleanup";
import { evaluateWorkbook } from "../lib/formula";
import type { CellValue, RowSpec, SheetSpec, WorkbookSpec } from "../lib/types";
import { applyScenario, goalSeek, goalSeekProblem, normalizeRef, scenarioResults } from "../lib/whatif";

let failed = 0;
function check(name: string, got: unknown, want: unknown) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(ok ? "ok  " : "FAIL", name, ok ? "" : `\n     got  ${JSON.stringify(got)}\n     want ${JSON.stringify(want)}`);
}
function near(name: string, got: number, want: number, tol = 1e-6) {
  const ok = Number.isFinite(got) && Math.abs(got - want) <= tol * Math.max(1, Math.abs(want));
  if (!ok) failed++;
  console.log(ok ? "ok  " : "FAIL", name, ok ? "" : `\n     got  ${got}\n     want ${want}`);
}

const row = (...xs: CellValue[]): RowSpec => ({ kind: "data", cells: xs.map((value) => ({ value })) });
const total = (...xs: CellValue[]): RowSpec => ({ kind: "total", cells: xs.map((value) => ({ value })) });
const sheetOf = (name: string, headers: string[], rows: RowSpec[]): SheetSpec => ({
  name,
  columns: headers.map((header) => ({ header, format: "text" })),
  rows,
});
const book = (...sheets: SheetSpec[]): WorkbookSpec => ({ title: "t", sheets });
const col = (s: SheetSpec, c: number) => s.rows.map((r) => r.cells[c]?.value ?? null);

// Scope
const s0 = sheetOf("S", ["A", "B", "C"], []);
check("scope single cell = whole sheet", cleanupScope(s0, { r1: 2, c1: 1, r2: 2, c2: 1 }), { cols: [0, 1, 2], whole: true });
check("scope columns", cleanupScope(s0, { r1: 1, c1: 1, r2: 5, c2: 2 }), { cols: [1, 2], whole: false });

// Duplicates
{
  const wb = book(
    sheetOf("Data", ["Name", "City", "Amt"], [
      row("Ann", "Oslo", 1),
      row("Bob", "Rome", 2),
      row(" ann", "oslo", 3),
      row("Bob", "Rome", 4),
      row(null, null, 5),
      row(null, null, 6),
      total("Total", null, "=SUM(C2:C7)"),
    ]),
  );
  check("dup rows (name+city)", duplicateRows(wb.sheets[0], [0, 1]), [2, 3]);
  check("dup rows all cols", duplicateRows(wb.sheets[0], [0, 1, 2]), []);
  const r = applyCleanup(wb, 0, { kind: "dedupe" }, [0, 1]);
  check("dedupe count", r.changed, 2);
  check("dedupe keeps first", col(wb.sheets[0], 2), [1, 2, 5, 6, "=SUM(C2:C5)"]);
  check("dedupe total sum", evaluateWorkbook(wb).get("0:6:3"), 14);
}

// Trim / case
check("trim", trimText("  New   York \t "), "New York");
check("title case", changeCaseText("hello WORLD-wide o'neil", "title"), "Hello World-Wide O'neil");
{
  const wb = book(sheetOf("S", ["T"], [row("  a  b "), row("=TRIM(A2)"), row(5), total(" x ")]));
  check("trim op count", applyCleanup(wb, 0, { kind: "trim" }, [0]).changed, 1);
  check("trim op leaves formulas/totals", col(wb.sheets[0], 0), ["a b", "=TRIM(A2)", 5, " x "]);
  check("upper op", (applyCleanup(wb, 0, { kind: "case", mode: "upper" }, [0]), col(wb.sheets[0], 0)), ["A B", "=TRIM(A2)", 5, " x "]);
}

// Split
{
  const wb = book(
    sheetOf("S", ["Name", "Score"], [row("Ann Lee", 1), row("Bo", 2), row("Cy de Vo", 3), total("T", "=SUM(B2:B4)")]),
    sheetOf("Other", ["X"], [row("=S!B2*2")]),
  );
  const r = applyCleanup(wb, 0, { kind: "split", column: 0, delimiter: " " }, [0]);
  check("split counts", [r.changed, r.addedColumns], [2, 2]);
  check("split headers", wb.sheets[0].columns.map((c) => c.header), ["Name", "Name 2", "Name 3", "Score"]);
  check("split row", wb.sheets[0].rows[2].cells.map((c) => c.value), ["Cy", "de", "Vo", 3]);
  check("split short row", wb.sheets[0].rows[1].cells.map((c) => c.value), ["Bo", null, null, 2]);
  check("split shifts formulas", [wb.sheets[0].rows[3].cells[3].value, wb.sheets[1].rows[0].cells[0].value], ["=SUM(D2:D4)", "=S!D2*2"]);
  const wb2 = book(sheetOf("S", ["P"], [row("1, 2;3"), row("x")]));
  applyCleanup(wb2, 0, { kind: "split", column: 0, delimiter: "," }, [0]);
  check("split comma -> numbers", wb2.sheets[0].rows[0].cells.map((c) => c.value), [1, "2;3"]);
}

// Dates
check("date iso", parseDateText("2024-03-05", false), "2024-03-05");
check("date us", parseDateText("03/05/2024", false), "2024-03-05");
check("date dayfirst", parseDateText("03/05/2024", true), "2024-05-03");
check("date unambiguous", parseDateText("25/12/2024", false), "2024-12-25");
check("date d mon y", parseDateText("5 Mar 2024", false), "2024-03-05");
check("date month d y", parseDateText("March 5 2024", false), "2024-03-05");
check("date month d, y ordinal", parseDateText("Tue, March 5th, 2024", false), "2024-03-05");
check("date 2-digit year", parseDateText("5-Mar-24", false), "2024-03-05");
check("date invalid day", parseDateText("2024-02-30", false), null);
check("date not a date", parseDateText("Mar", false), null);
{
  const wb = book(sheetOf("S", ["D", "N"], [row("5 Mar 2024", "x"), row("2024-3-6", null), row("soon", null), row("=TODAY()", null)]));
  const r = applyCleanup(wb, 0, { kind: "dates", dayFirst: false }, [0]);
  check("dates op", [r.changed, col(wb.sheets[0], 0), wb.sheets[0].columns[0].format], [2, ["2024-03-05", "2024-03-06", "soon", "=TODAY()"], "date"]);
  check("dates evaluate as serials", evaluateWorkbook(book({ ...wb.sheets[0], rows: [...wb.sheets[0].rows, row("=A3-A2")] })).get("0:6:1"), 1);
}

// Numbers
check("num comma", parseNumberText("1,200"), { value: 1200 });
check("num dollar", parseNumberText("$5"), { value: 5, format: "currency", currency: "USD" });
check("num euro neg", parseNumberText("-€3.50"), { value: -3.5, format: "currency", currency: "EUR" });
check("num parens", parseNumberText("(500)"), { value: -500 });
check("num percent", parseNumberText("12.5%"), { value: 0.125, format: "percent" });
check("num zip kept", parseNumberText("01234"), null);
check("num bad grouping", parseNumberText("1,20"), null);
check("num text", parseNumberText("abc"), null);
{
  const wb = book(sheetOf("S", ["A"], [row("1,200"), row("$5"), row("n/a"), row("=1+1"), row(7)]));
  const r = applyCleanup(wb, 0, { kind: "numbers" }, [0]);
  check("numbers op", [r.changed, col(wb.sheets[0], 0), wb.sheets[0].rows[1].cells[0].format], [2, [1200, 5, "n/a", "=1+1", 7], "currency"]);
}

// Fill blanks
{
  const mk = () =>
    book(sheetOf("S", ["Region", "Amt"], [row("East", 1), row(null, 2), row("", 3), total(null, null), row(null, 4), row("West", "=B6*2"), row(null, null)]));
  let wb = mk();
  const r = applyCleanup(wb, 0, { kind: "fill", mode: "above" }, [0, 1]);
  check("fill above count", r.changed, 4);
  check("fill above col A", col(wb.sheets[0], 0), ["East", "East", "East", null, null, "West", "West"]);
  check("fill above formula moves", wb.sheets[0].rows[6].cells[1].value, "=B7*2");
  wb = mk();
  applyCleanup(wb, 0, { kind: "fill", mode: "zero" }, [1]);
  check("fill zero", col(wb.sheets[0], 1), [1, 2, 3, null, 4, "=B6*2", 0]);
  wb = mk();
  applyCleanup(wb, 0, { kind: "fill", mode: "value", value: "n/a" }, [0]);
  check("fill value", col(wb.sheets[0], 0), ["East", "n/a", "n/a", null, "n/a", "West", "n/a"]);
}

// Scenarios
{
  const wb = book(
    sheetOf("Inputs", ["Item", "Value"], [row("Price", 10), row("Units", 100), row("Cost", 400)]),
    sheetOf("Income", ["Line", "Amount"], [row("Revenue", "=Inputs!B2*Inputs!B3"), row("Profit", "=B2-Inputs!B4")]),
  );
  check("normalize ref", [normalizeRef(wb, "income!b3"), normalizeRef(wb, "B2", 0), normalizeRef(wb, "Nope!A2"), normalizeRef(wb, "A1")], [
    "Income!B3",
    "Inputs!B2",
    null,
    null,
  ]);
  const scenarios = [
    { name: "Best", values: { "Inputs!B2": 12, "Inputs!B3": 150 } },
    { name: "Worst", values: { "Inputs!B2": 8, "Inputs!B3": 60 } },
  ];
  check("scenario results", scenarioResults(wb, scenarios, ["Income!B3", "Income!B2"]), [
    [1400, 1800],
    [80, 480],
  ]);
  check("scenario leaves original", wb.sheets[0].rows[0].cells[1].value, 10);
  applyScenario(wb, scenarios[1]);
  check("apply scenario", [wb.sheets[0].rows[0].cells[1].value, wb.sheets[0].rows[1].cells[1].value], [8, 60]);

  // Goal seek
  const fresh = book(
    sheetOf("Inputs", ["Item", "Value"], [row("Price", 10), row("Units", 100), row("Cost", 400)]),
    sheetOf("Income", ["Line", "Amount"], [row("Revenue", "=Inputs!B2*Inputs!B3"), row("Profit", "=B2-Inputs!B4"), row("Sq", "=Inputs!B2^2"), row("Step", "=IF(Inputs!B2>5,10,0)")]),
  );
  const g = goalSeek(fresh, "Income!B3", 2000, "Inputs!B2");
  check("goal seek linear found", [g.found, g.value], [true, 24]);
  check("goal seek leaves original", fresh.sheets[0].rows[0].cells[1].value, 10);
  const q = goalSeek(fresh, "Income!B4", 2, "Inputs!B2");
  near("goal seek sqrt(2)", q.value, Math.SQRT2);
  check("goal seek sqrt found", q.found, true);
  check("goal seek no solution (x^2 = -1)", goalSeek(fresh, "Income!B4", -1, "Inputs!B2").found, false);
  check("goal seek no solution (step)", goalSeek(fresh, "Income!B5", 5, "Inputs!B2").found, false);
  const big = goalSeek(fresh, "Income!B2", 1e7, "Inputs!B3");
  check("goal seek big target", [big.found, big.value], [true, 1e6]);
  check("goal seek needs formula", goalSeekProblem(fresh, "Inputs!B3", "Inputs!B2"), "The result cell must contain a formula.");
  check("goal seek needs constant", goalSeekProblem(fresh, "Income!B3", "Income!B2"), "The changing cell must hold a number, not a formula.");
  const loan = book(sheetOf("L", ["K", "V"], [row("Rate", 0.05), row("Payment", "=-PMT(B2/12,360,300000)")]));
  const lr = goalSeek(loan, "L!B3", 1500, "L!B2");
  check("goal seek loan rate found", lr.found, true);
  near("goal seek loan payment", (evaluateWorkbook({ ...loan, sheets: [{ ...loan.sheets[0], rows: [row("Rate", lr.value), loan.sheets[0].rows[1]] }] }).get("0:3:2") as number), 1500, 1e-6);
}

console.log(failed ? `\n${failed} check(s) failed` : "\nall checks passed");
process.exit(failed ? 1 : 0);
