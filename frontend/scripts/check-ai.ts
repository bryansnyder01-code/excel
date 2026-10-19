// Checks for the =AI() formula cache/queue and for live-data sheet writing:  npx tsx scripts/check-ai.ts
import { __aiTest, aiKey, flush, AI_LOADING } from "../lib/aiFormula";
import { evaluateCached, evaluateWorkbook } from "../lib/formula";
import { detectKind, isDue, toCellValue, writeTable } from "../lib/connections";
import { useStore } from "../lib/store";
import { scheduleToLocal, scheduleToUtc, type AiEvalItem, type Schedule } from "../lib/aiApi";
import type { WorkbookSpec } from "../lib/types";

let failed = 0;
function check(name: string, got: unknown, want: unknown) {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  if (!ok) failed++;
  console.log(ok ? "ok  " : "FAIL", name, ok ? "" : `=> got ${JSON.stringify(got)}, want ${JSON.stringify(want)}`);
}
const tick = () => new Promise((r) => setTimeout(r, 5));
const v = (x: unknown) => ({ value: x as never });

const calls: AiEvalItem[][] = [];
let reply: (items: AiEvalItem[]) => Record<string, string> = (items) =>
  Object.fromEntries(items.map((i) => [i.key, `${i.prompt}:${i.inputs.join("|")}`]));
let fail: string | null = null;
__aiTest.setSender(async (items) => {
  calls.push(items);
  if (fail) throw new Error(fail);
  return reply(items);
});
let clock = 1_000_000;
__aiTest.setClock(() => clock);

function book(): WorkbookSpec {
  return {
    title: "t",
    sheets: [
      {
        name: "Spend",
        columns: [{ header: "Item" }, { header: "Amount", format: "currency" }, { header: "Category" }, { header: "Date", format: "date" }],
        rows: [
          { cells: [v("Coffee"), v(4.5), v('=AI("categorise", A2, B2)'), v("2024-03-05")] },
          { cells: [v("Taxi"), v(20), v('=AI("categorise", A3, B3)'), v("2024-03-06")] },
          { cells: [v("Coffee"), v(4.5), v('=AI("categorise", A4, B4)'), v(null)] },
          { cells: [v("Rent"), v(900), { value: '=AI("categorise", A5, B5)', ai_value: "Housing" }, v(null)] },
          { cells: [v(null), v(null), v('=AI("when", D2)'), v(null)] },
          { cells: [v(null), v(null), v('=AI("list", A2:A3)'), v(null)] },
        ],
      },
    ],
  };
}

const at = (wb: WorkbookSpec, row: number) => evaluateCached(wb).get(`0:${row}:3`);
const cur = () => useStore.getState().workbook;

async function main() {
  // Keys are the prompt plus inputs as displayed (currency and dates formatted).
  check("key", aiKey("categorise", ["Coffee", "$4.50"]), '["categorise","Coffee","$4.50"]');

  // A copy that isn't the open workbook never queues requests.
  __aiTest.reset();
  const copy = book();
  check("non-current shows loading", at(copy, 2), AI_LOADING);
  check("non-current trusts ai_value", at(copy, 5), "Housing");
  check("non-current queues nothing", __aiTest.state().queued, []);

  // The open workbook queues each distinct question once.
  __aiTest.reset();
  useStore.setState({ workbook: book(), past: [], future: [] });
  const first = cur();
  check("loading while fetched", [at(first, 2), at(first, 3), at(first, 4)], [AI_LOADING, AI_LOADING, AI_LOADING]);
  check("saved ai_value shown", at(first, 5), "Housing");
  check("duplicate inputs queued once", __aiTest.state().queued.length, 4);
  check("date input displayed", __aiTest.state().queued.includes(aiKey("when", ["2024-03-05"])), true);
  check("range input joined", __aiTest.state().queued.includes(aiKey("list", ["Coffee, Taxi"])), true);
  check("currency input displayed", __aiTest.state().queued.includes(aiKey("categorise", ["Coffee", "$4.50"])), true);

  // Evaluating again (new object, same content) doesn't queue twice.
  useStore.setState({ workbook: { ...first } });
  at(cur(), 2);
  check("no re-queue", __aiTest.state().queued.length, 4);

  await flush();
  check("one batched call", calls.length, 1);
  check("batch size", calls[0].length, 4);
  check("no undo step", useStore.getState().past.length, 0);
  const after = cur();
  check("workbook replaced to re-evaluate", after !== first, true);
  check("answer shown", at(after, 2), "categorise:Coffee|$4.50");
  check("same key shares answer", at(after, 4), "categorise:Coffee|$4.50");
  await tick();
  const written = cur();
  check("ai_value written", written.sheets[0].rows.map((r) => r.cells[2].ai_value ?? null), [
    "categorise:Coffee|$4.50",
    "categorise:Taxi|$20.00",
    "categorise:Coffee|$4.50",
    "Housing",
    "when:2024-03-05",
    "list:Coffee, Taxi",
  ]);
  check("copy-on-write keeps untouched rows", written.sheets[0].rows[3] === after.sheets[0].rows[3], true);
  check("still no undo step", useStore.getState().past.length, 0);
  await tick();
  check("settles (no endless writes)", cur() === written, true);

  // A formula copied with its ai_value (fill down, paste) to new inputs isn't trusted.
  useStore.getState().mutate((wb) => {
    wb.sheets[0].rows.push({ cells: [v("Bus"), v(3), { value: '=AI("categorise", A8, B8)', ai_value: "categorise:Coffee|$4.50" }] });
  });
  check("copied ai_value not trusted", at(cur(), 8), AI_LOADING);
  __aiTest.reset();
  useStore.getState().mutate((wb) => {
    wb.sheets[0].rows.pop();
  });
  useStore.setState({ workbook: structuredClone(cur()) });
  // After a reset (like a page load) the saved answers are trusted again.
  check("trusted on load", [at(cur(), 2), at(cur(), 5)], ["categorise:Coffee|$4.50", "Housing"]);
  check("nothing queued on load", __aiTest.state().queued, []);

  // Changing an input asks again; a saved ai_value isn't trusted for new inputs.
  useStore.getState().mutate((wb) => {
    wb.sheets[0].rows[3].cells[0].value = "Mortgage";
  });
  check("changed input loads", at(cur(), 5), AI_LOADING);
  check("changed input queued", __aiTest.state().queued, [aiKey("categorise", ["Mortgage", "$900.00"])]);

  // Failures show #AI! and aren't asked again for a minute.
  fail = "You've used today's AI allowance.";
  await flush();
  check("error shown", at(cur(), 5), { error: "#AI!" });
  useStore.setState({ workbook: { ...cur() } });
  at(cur(), 5);
  check("failed key not re-queued", __aiTest.state().queued, []);
  clock += 61_000;
  fail = null;
  useStore.setState({ workbook: { ...cur() } });
  check("retried after a minute", at(cur(), 5), AI_LOADING);
  check("re-queued", __aiTest.state().queued.length, 1);
  // In-flight keys aren't queued again.
  reply = (items) => Object.fromEntries(items.map((i) => [i.key, "42"]));
  const pending = flush();
  useStore.setState({ workbook: { ...cur() } });
  at(cur(), 5);
  check("in-flight not re-queued", __aiTest.state().queued, []);
  await pending;
  check("numeric answers are numbers", at(cur(), 5), 42);
  await tick();
  check("numeric ai_value saved as text", cur().sheets[0].rows[3].cells[2].ai_value, "42");

  // Batches are capped at 100 items.
  __aiTest.reset();
  calls.length = 0;
  reply = (items) => Object.fromEntries(items.map((i) => [i.key, "x"]));
  const big: WorkbookSpec = {
    title: "b",
    sheets: [{ name: "S", columns: [{ header: "N" }, { header: "AI" }], rows: Array.from({ length: 150 }, (_, i) => ({ cells: [v(i), v(`=AI("double", A${i + 2})`)] })) }],
  };
  useStore.setState({ workbook: big });
  evaluateCached(big);
  await flush();
  await flush();
  check("batches of 100", calls.map((c) => c.length), [100, 50]);

  // Plain evaluation still works with AI in other formulas.
  check("AI inside other formulas", evaluateWorkbook({ title: "x", sheets: [{ name: "S", columns: [{ header: "a" }], rows: [{ cells: [v('=LEN(AI("x"))')] }] }] }).get("0:2:1"), AI_LOADING.length);

  // Connections: writing a fetched table into a sheet.
  check("detect gsheet", detectKind("https://docs.google.com/spreadsheets/d/abc/edit#gid=0"), "gsheet");
  check("detect json", detectKind("https://api.example.com/v1/items"), "json");
  check("detect json ext", detectKind("https://x.com/data.json"), "json");
  check("detect csv", detectKind("https://x.com/export.csv"), "csv");
  check("numbers as numbers", [toCellValue("12"), toCellValue("-3.5"), toCellValue("00123"), toCellValue("abc"), toCellValue(""), toCellValue(7)], [12, -3.5, "00123", "abc", null, 7]);

  const wb: WorkbookSpec = {
    title: "c",
    sheets: [
      { name: "Sheet1", columns: [{ header: "A" }], rows: [] },
      {
        name: "Sales",
        columns: [{ header: "Region", format: "text", width: 20 }, { header: "Revenue", format: "currency", currency: "EUR" }, { header: "Old" }],
        rows: [{ kind: "data", cells: [v("x"), v(1), v(2)] }, { kind: "total", cells: [v("Total"), v("=SUM(B2:B2)"), v(null)] }],
      },
    ],
  };
  const res = writeTable(wb, "sales", { columns: ["revenue", "Region", "Units"], rows: [["100.5", "East", 3], [null, "West", "4"]] });
  check("existing sheet found by name", res, { index: 1, created: false });
  check("columns replaced, formats kept by header", wb.sheets[1].columns, [
    { header: "revenue", format: "currency", currency: "EUR" },
    { header: "Region", format: "text", width: 20 },
    { header: "Units", format: "text" },
  ]);
  check("rows replaced", wb.sheets[1].rows.map((r) => [r.kind, ...r.cells.map((c) => c.value)]), [
    ["data", 100.5, "East", 3],
    ["data", null, "West", 4],
  ]);
  const made = writeTable(wb, "Live: feed", { columns: ["id"], rows: [[1]] });
  check("missing sheet created", [made, wb.sheets[2].name, wb.sheets[2].rows.length], [{ index: 2, created: true }, "Live  feed".replace(/\s+/g, " "), 1]);
  check("due when never refreshed", isDue({ id: "1", name: "n", kind: "csv", url: "u", sheet: "s", refresh_minutes: 15 }), true);
  check("not due when manual", isDue({ id: "1", name: "n", kind: "csv", url: "u", sheet: "s", refresh_minutes: null }), false);
  check("seconds timestamps", isDue({ id: "1", name: "n", kind: "csv", url: "u", sheet: "s", refresh_minutes: 15, last_refreshed: Math.floor(Date.now() / 1000) - 60 }), false);

  // Report schedules: local choices survive the trip to UTC and back.
  const trips: Schedule[] = [
    { frequency: "daily", weekday: null, day: null, hour: 0 },
    { frequency: "weekly", weekday: 6, day: null, hour: 23 },
    { frequency: "weekly", weekday: 0, day: null, hour: 1 },
    { frequency: "monthly", weekday: null, day: 15, hour: 8 },
  ];
  check("schedule round trip", trips.map((t) => scheduleToLocal(scheduleToUtc(t))), trips);

  console.log(failed ? `\n${failed} check(s) failed` : "\nAll AI and connection checks passed");
  process.exit(failed ? 1 : 0);
}

main();
