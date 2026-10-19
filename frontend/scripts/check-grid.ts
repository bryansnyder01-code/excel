// Checks for the grid helpers:  npx tsx scripts/check-grid.ts
import { diffWorkbooks } from "../lib/diff";
import { fillSeries, parseTsv, toTsv } from "../lib/fill";
import { formatValue, parseInput, typedFormat } from "../lib/format";
import { formulaContext, suggestFunctions } from "../lib/functions";
import { formulaRefs, offsetFormula } from "../lib/refs";

let failed = 0;
function check(name: string, got: unknown, want: unknown) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(ok ? "ok  " : "FAIL", name, ok ? "" : `\n     got  ${JSON.stringify(got)}\n     want ${JSON.stringify(want)}`);
}
const vals = (cells: { value?: unknown }[]) => cells.map((c) => c.value);
const v = (...xs: unknown[]) => xs.map((x) => ({ value: x as never }));

// Clipboard text
check("tsv simple", parseTsv("a\tb\n1\t2\n"), [["a", "b"], ["1", "2"]]);
check("tsv quoted newline", parseTsv('"x\ny"\tz'), [["x\ny", "z"]]);
check("tsv round trip", parseTsv(toTsv([["a\tb", 'q"t'], ["", "3"]])), [["a\tb", 'q"t'], ["", "3"]]);
check("tsv windows", parseTsv("1\t2\r\n3\t4\r\n"), [["1", "2"], ["3", "4"]]);

// Moving formulas
check("offset relative", offsetFormula("=A1+$B$2+C$3+$D4", 2, 1), "=B3+$B$2+D$3+$D6");
check("offset range + sheet", offsetFormula("=SUM('My Data'!B2:B10)", 1, 0), "=SUM('My Data'!B3:B11)");
check("offset strings untouched", offsetFormula('=IF(A2="A2",1,0)', 1, 0), '=IF(A3="A2",1,0)');
check("offset off sheet", offsetFormula("=A1", -1, 0), "=#REF!");
check("refs", formulaRefs("=SUM(B2:B5)/C1+'X'!A1").map((r) => [r.sheet ?? "", r.r1, r.c1, r.r2, r.c2]), [["", 2, 2, 5, 2], ["", 1, 3, 1, 3], ["X", 1, 1, 1, 1]]);

// Fill handle
check("fill numbers", vals(fillSeries(v(1, 3), 3, "row", 1)), [5, 7, 9]);
check("fill single number repeats", vals(fillSeries(v(4), 2, "row", 1)), [4, 4]);
check("fill up", vals(fillSeries(v(3, 2), 2, "row", -1)), [1, 0]);
check("fill dates", vals(fillSeries(v("2026-01-30"), 3, "row", 1)), ["2026-01-31", "2026-02-01", "2026-02-02"]);
check("fill weekly dates", vals(fillSeries(v("2026-01-01", "2026-01-08"), 1, "row", 1)), ["2026-01-15"]);
check("fill months", vals(fillSeries(v("Nov"), 3, "col", 1)), ["Dec", "Jan", "Feb"]);
check("fill weekdays", vals(fillSeries(v("Monday", "Wednesday"), 2, "row", 1)), ["Friday", "Sunday"]);
check("fill quarters", vals(fillSeries(v("Q3"), 2, "row", 1)), ["Q4", "Q1"]);
check("fill numbered text", vals(fillSeries(v("Item 9"), 2, "row", 1)), ["Item 10", "Item 11"]);
check("fill formula", vals(fillSeries(v("=B2*C2"), 2, "row", 1)), ["=B3*C3", "=B4*C4"]);
check("fill formula right", vals(fillSeries(v("=SUM(B2:B9)"), 2, "col", 1)), ["=SUM(C2:C9)", "=SUM(D2:D9)"]);
check("fill formula pattern", vals(fillSeries(v("=A1", "=A1*2"), 2, "row", 1)), ["=A3", "=A3*2"]);
check("fill keeps style", fillSeries([{ value: 1, bold: true }, { value: 2 }], 1, "row", 1)[0], { bold: true, value: 3 });
check("fill text repeats", vals(fillSeries(v("East", "West"), 3, "row", 1)), ["East", "West", "East"]);

// Formula autocomplete
const ctx = (s: string) => formulaContext(s, s.length);
check("partial", ctx("=SU").partial, { text: "SU", start: 1 });
check("cell ref is not a function", ctx("=B1").partial, null);
check("inside call", ctx("=IF(A1>0,").call, { name: "IF", arg: 1 });
check("nested call", ctx("=SUM(ROUND(A1,").call, { name: "ROUND", arg: 1 });
check("string ignored", ctx('=IF(A1="x,y",').call, { name: "IF", arg: 1 });
check("suggest", suggestFunctions("SUMI").map((f) => f.name), ["SUMIF", "SUMIFS"]);

// Number formats
check("euro", formatValue(1234.5, { fmt: "currency", currency: "EUR" }).text, "€1,234.50");
check("yen", formatValue(1234.5, { fmt: "currency", currency: "JPY" }).text, "¥1,235");
check("negative pound", formatValue(-5, { fmt: "currency", currency: "GBP", decimals: 0 }).text, "-£5");
check("decimals", formatValue(2.5, { fmt: "decimal", decimals: 3 }).text, "2.500");
check("percent decimals", formatValue(0.1234, { fmt: "percent", decimals: 2 }).text, "12.34%");

check("typed percent", typedFormat("15%"), { format: "percent" });
check("typed euro", typedFormat("€1,200.50"), { format: "currency", currency: "EUR" });
check("typed CA$", typedFormat("CA$9"), { format: "currency", currency: "CAD" });
check("typed plain", typedFormat("1200"), null);
check("parse CA$", parseInput("CA$1,234.5"), 1234.5);
check("parse euro", parseInput("€-3"), -3);

// Review diff
const before = { title: "t", sheets: [{ name: "A", columns: [{ header: "X" }], rows: [{ cells: v(1, 2) }] }] };
const after = {
  title: "t",
  sheets: [
    { name: "A", columns: [{ header: "X" }], rows: [{ cells: [{ value: 1 }, { value: 3, bold: true }] }] },
    { name: "B", columns: [{ header: "Y" }], rows: [] },
  ],
};
check(
  "diff",
  diffWorkbooks(before as never, after as never).map((d) => [d.name, d.status, [...d.cells]]),
  [["A", "changed", ["1:1"]], ["B", "added", ["0:0"]]],
);

console.log(failed ? `\n${failed} failed` : "\nall grid checks passed");
process.exit(failed ? 1 : 0);
