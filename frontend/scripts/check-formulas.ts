// Smoke test for the live formula engine:  npx tsx scripts/check-formulas.ts
import fs from "fs";
import path from "path";
import { evaluateWorkbook } from "../lib/formula";
const v = (x?: unknown) => ({ value: (x ?? null) as never });
const wb = {
  title: "t",
  sheets: [{
    name: "Data",
    columns: [{ header: "Region" }, { header: "Q" }, { header: "Rev", format: "currency" }, { header: "Out" }],
    rows: [
      { cells: [v("East"), v("Q1"), v(100), v("=MAX($C$2:$C$6)")] },
      { cells: [v("West"), v("Q1"), v(80), v("=SUMIFS($C$2:$C$6,$A$2:$A$6,\"East\")")] },
      { cells: [v("East"), v("Q2"), v(120), v("=IFERROR(AVERAGEIFS($C$2:$C$6,$A$2:$A$6,A2),0)")] },
      { cells: [v("North"), v("Q2"), v(50), v("=COUNTIFS($A$2:$A$6,\"East\",$B$2:$B$6,\"Q2\")")] },
      { cells: [v("West"), v("Q2"), v(90), v("=-PMT(0.065/12,60,25000)")] },
      { cells: [v(), v(), v(), v("=MINIFS(C2:C6,A2:A6,\"West\")")] },
      { cells: [v(), v(), v(), v("=COUNTA(A2:A8)")] },
      { cells: [v(), v(), v(), v("=IF(C2=MAX($C$2:$C$6),\"Top\",\"\")")] },
      { cells: [v(), v(), v(), v("=XLOOKUP(\"North\",A2:A6,C2:C6)")] },
      { cells: [v(), v(), v(), v("=SUM(C2:C6)/0")] },
      { cells: [v(), v(), v(), v("=MAX(C:C)")] },
    ],
  }],
};
const expected: unknown[] = [120, 220, 110, 1, 489.15, 80, 5, "", 50, { error: "#DIV/0!" }, 120];
const out = evaluateWorkbook(wb as never);
let failed = 0;
expected.forEach((want, i) => {
  const got = out.get(`0:${i + 2}:4`);
  const ok = typeof want === "number" ? Math.abs((got as number) - want) < 0.01 : JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(ok ? "ok  " : "FAIL", wb.sheets[0].rows[i].cells[3].value, "=>", JSON.stringify(got));
});

// Every starter template should evaluate without a single error value.
const templates = path.join(__dirname, "../../backend/app/templates");
if (fs.existsSync(templates)) {
  for (const f of fs.readdirSync(templates)) {
    const t = JSON.parse(fs.readFileSync(path.join(templates, f), "utf8"));
    const errs = [...evaluateWorkbook(t.workbook).entries()].filter(([, x]) => x && typeof x === "object");
    if (errs.length) failed++;
    console.log(errs.length ? "FAIL" : "ok  ", `template ${f}`, errs.length ? JSON.stringify(errs.slice(0, 3)) : "");
  }
}
process.exit(failed ? 1 : 0);
