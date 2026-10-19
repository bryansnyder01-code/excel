"""Regenerate the starter templates in app/templates/.

    python -m scripts.make_templates

Each template is a normal WorkbookSpec with live formulas, so it opens in the grid
and can be tailored by chat like anything the AI generates.
"""

import json
from pathlib import Path

from app.schemas import WorkbookSpec

OUT = Path(__file__).resolve().parent.parent / "app" / "templates"
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def c(v=None, **k):
    return {"value": v, **k}


def col(header, fmt="text", width=None):
    d = {"header": header, "format": fmt}
    if width:
        d["width"] = width
    return d


def budget():
    months = MONTHS[:6]
    inc = [("Product sales", 12000, 13500, 12800, 14200, 15100, 16000),
           ("Consulting", 4000, 3500, 5200, 4800, 5000, 6100),
           ("Subscriptions", 1800, 1900, 2050, 2100, 2300, 2400)]
    exp = [("Rent", *[3000] * 6), ("Payroll", 8500, 8500, 9000, 9000, 9200, 9200),
           ("Marketing", 1200, 1500, 1100, 1800, 2000, 1700), ("Software", 450, 450, 480, 480, 500, 500),
           ("Utilities", 320, 300, 290, 280, 310, 330)]

    def block(data):
        rows = [{"cells": [c(n)] + [c(v) for v in vals] + [c(f"=SUM(B{i + 2}:G{i + 2})")]} for i, (n, *vals) in enumerate(data)]
        last = len(data) + 1
        rows.append({"kind": "total", "cells": [c("Total")] + [c(f"=SUM({L}2:{L}{last})") for L in "BCDEFGH"]})
        return rows

    summary = [{"cells": [c(m), c(f"='Income'!{L}5"), c(f"='Expenses'!{L}7"), c(f"=B{i + 2}-C{i + 2}"),
                          c(f"=IFERROR(D{i + 2}/B{i + 2},0)")]} for i, (m, L) in enumerate(zip(months, "BCDEFG"))]
    summary.append({"kind": "total", "cells": [c("Total"), c("=SUM(B2:B7)"), c("=SUM(C2:C7)"), c("=B8-C8"), c("=IFERROR(D8/B8,0)")]})
    money = [col(m, "currency") for m in months] + [col("Total", "currency")]
    return {
        "id": "small-business-budget", "title": "Small business budget", "category": "Finance",
        "description": "Six months of income and expenses with a net profit summary and chart.",
        "workbook": {"title": "Small Business Budget", "sheets": [
            {"name": "Income", "columns": [col("Source")] + money, "rows": block(inc)},
            {"name": "Expenses", "columns": [col("Category")] + money, "rows": block(exp)},
            {"name": "Summary", "columns": [col("Month"), col("Income", "currency"), col("Expenses", "currency"),
                                            col("Net Profit", "currency"), col("Margin", "percent")],
             "rows": summary,
             "conditional_formats": [{"range": "D2:D8", "operator": "negative", "style": "bad"}],
             "charts": [{"type": "column", "title": "Income vs expenses", "categories": "A2:A7",
                         "series": [{"name": "Income", "range": "B2:B7"}, {"name": "Expenses", "range": "C2:C7"}]}]},
        ]},
    }


def invoice():
    items = [("Website design", 1, 2400), ("Hosting (12 months)", 12, 25), ("Logo package", 1, 650), ("Copywriting (hours)", 8, 85)]
    rows = [{"cells": [c(d), c(q), c(p), c(f"=B{i + 2}*C{i + 2}")]} for i, (d, q, p) in enumerate(items)]
    n = len(items) + 1
    rows += [
        {"kind": "subtotal", "cells": [c("Subtotal"), c(), c(), c(f"=SUM(D2:D{n})")]},
        {"cells": [c("Tax rate"), c(), c(0.08, format="percent"), c(f"=D{n + 1}*C{n + 2}")]},
        {"cells": [c("Discount"), c(), c(), c(0)]},
        {"kind": "total", "cells": [c("Amount due"), c(), c(), c(f"=D{n + 1}+D{n + 2}-D{n + 3}")]},
    ]
    return {
        "id": "invoice", "title": "Invoice", "category": "Business",
        "description": "Line items with quantity x price, subtotal, tax and amount due.",
        "workbook": {"title": "Invoice", "sheets": [
            {"name": "Invoice", "show_gridlines": False,
             "columns": [col("Description", width=34), col("Qty", "integer"), col("Unit price", "currency"), col("Amount", "currency")],
             "rows": rows},
            {"name": "Details", "columns": [col("Field", width=18), col("Value", width=36)], "rows": [
                {"cells": [c("Invoice #"), c("INV-1001")]}, {"cells": [c("Date"), c("2026-10-18", format="date")]},
                {"cells": [c("Due date"), c("=B3+30", format="date")]}, {"cells": [c("Bill to"), c("Client name")]},
                {"cells": [c("Status"), c("Unpaid")]}],
             "validations": [{"range": "B6", "type": "list", "options": ["Unpaid", "Paid", "Overdue"]}]},
        ]},
    }


def loan():
    n = 60
    rows = []
    for i in range(n):
        r = i + 2
        bal = "='Loan'!B2" if i == 0 else f"=F{r - 1}"
        rows.append({"cells": [c(i + 1), c(bal), c(f"=-PMT('Loan'!$B$3/12,'Loan'!$B$4*12,'Loan'!$B$2)"),
                               c(f"=B{r}*'Loan'!$B$3/12"), c(f"=C{r}-D{r}"), c(f"=MAX(B{r}-E{r},0)")]})
    rows.append({"kind": "total", "cells": [c("Total"), c(), c(f"=SUM(C2:C{n + 1})"), c(f"=SUM(D2:D{n + 1})"), c(f"=SUM(E2:E{n + 1})"), c()]})
    return {
        "id": "loan-amortization", "title": "Loan amortization", "category": "Finance",
        "description": "Monthly payment, interest and principal schedule for a 5-year loan.",
        "workbook": {"title": "Loan Amortization", "sheets": [
            {"name": "Loan", "columns": [col("Input", width=22), col("Value", "currency")], "rows": [
                {"cells": [c("Loan amount"), c(25000)]},
                {"cells": [c("Annual interest rate"), c(0.065, format="percent")]},
                {"cells": [c("Term (years)"), c(5, format="integer")]},
                {"kind": "total", "cells": [c("Monthly payment"), c("=-PMT(B3/12,B4*12,B2)")]},
                {"cells": [c("Total interest"), c("='Schedule'!D62")]}]},
            {"name": "Schedule", "columns": [col("Month", "integer"), col("Opening balance", "currency"), col("Payment", "currency"),
                                             col("Interest", "currency"), col("Principal", "currency"), col("Closing balance", "currency")],
             "rows": rows,
             "charts": [{"type": "line", "title": "Remaining balance", "categories": "A2:A61",
                         "series": [{"name": "Closing balance", "range": "F2:F61"}]}]},
        ]},
    }


def inventory():
    items = [("SKU-001", "Wireless mouse", "Electronics", 42, 15, 18.5), ("SKU-002", "USB-C cable", "Electronics", 8, 25, 6.2),
             ("SKU-003", "Notebook A5", "Stationery", 120, 40, 2.1), ("SKU-004", "Desk lamp", "Furniture", 5, 10, 34.0),
             ("SKU-005", "Stapler", "Stationery", 18, 10, 7.5), ("SKU-006", "Monitor stand", "Furniture", 11, 8, 45.0),
             ("SKU-007", "Webcam", "Electronics", 3, 6, 59.0)]
    rows = [{"cells": [c(s), c(n), c(cat), c(q), c(rp), c(cost), c(f"=D{i + 2}*F{i + 2}"),
                       c(f'=IF(D{i + 2}<=E{i + 2},"Reorder","OK")')]} for i, (s, n, cat, q, rp, cost) in enumerate(items)]
    last = len(items) + 1
    rows.append({"kind": "total", "cells": [c("Total"), c(), c(), c(f"=SUM(D2:D{last})"), c(), c(), c(f"=SUM(G2:G{last})"),
                                            c(f'=COUNTIF(H2:H{last},"Reorder")&" to reorder"')]})
    return {
        "id": "inventory", "title": "Inventory tracker", "category": "Operations",
        "description": "Stock levels, reorder points, stock value and a reorder flag.",
        "workbook": {"title": "Inventory Tracker", "sheets": [
            {"name": "Inventory", "columns": [col("SKU"), col("Item", width=22), col("Category"), col("In stock", "integer"),
                                              col("Reorder at", "integer"), col("Unit cost", "currency"), col("Stock value", "currency"), col("Status")],
             "rows": rows,
             "validations": [{"range": "C2:C200", "type": "list", "options": ["Electronics", "Stationery", "Furniture", "Other"]},
                             {"range": "D2:E200", "type": "whole", "min": 0, "message": "Enter a whole number of 0 or more."}]},
        ]},
    }


def timesheet():
    days = [("2026-10-12", "Mon", "09:00", "17:30", 0.5), ("2026-10-13", "Tue", "08:45", "17:15", 0.5),
            ("2026-10-14", "Wed", "09:00", "18:00", 1), ("2026-10-15", "Thu", "09:15", "17:00", 0.5),
            ("2026-10-16", "Fri", "09:00", "16:00", 0.5)]
    rows = []
    for i, (d, day, start, end, brk) in enumerate(days):
        r = i + 2
        sh, sm = map(int, start.split(":"))
        eh, em = map(int, end.split(":"))
        rows.append({"cells": [c(d), c(day), c(sh + sm / 60), c(eh + em / 60), c(brk), c(f"=MAX(D{r}-C{r}-E{r},0)"),
                               c(f"=MAX(F{r}-8,0)")]})
    rows.append({"kind": "total", "cells": [c("Total"), c(), c(), c(), c("=SUM(E2:E6)"), c("=SUM(F2:F6)"), c("=SUM(G2:G6)")]})
    rows.append({"cells": [c("Hourly rate"), c(), c(), c(), c(), c(32, format="currency")]})
    rows.append({"kind": "total", "cells": [c("Gross pay"), c(), c(), c(), c(), c("=(F7-G7)*F8+G7*F8*1.5", format="currency")]})
    return {
        "id": "timesheet", "title": "Weekly timesheet", "category": "HR",
        "description": "Daily hours with breaks, overtime over 8 hours, and gross pay.",
        "workbook": {"title": "Weekly Timesheet", "sheets": [
            {"name": "Timesheet", "columns": [col("Date", "date"), col("Day"), col("Start (h)", "decimal"), col("End (h)", "decimal"),
                                              col("Break (h)", "decimal"), col("Hours", "decimal"), col("Overtime", "decimal")],
             "rows": rows},
        ]},
    }


def sales():
    reps = [("Ava Chen", 42000, 51000, 47500, 60200), ("Liam Patel", 38500, 36000, 44100, 48800),
            ("Sofia Garcia", 51200, 49800, 55300, 58100), ("Noah Kim", 29900, 34500, 31200, 39700),
            ("Mia Johnson", 45600, 47200, 52900, 50300)]
    rows = []
    for i, (name, *q) in enumerate(reps):
        r = i + 2
        rows.append({"cells": [c(name)] + [c(v) for v in q] + [c(f"=SUM(B{r}:E{r})"), c(f"=AVERAGE(B{r}:E{r})"),
                                                                c(f'=IF(F{r}=MAX($F$2:$F$6),"Top performer","")')]})
    rows.append({"kind": "total", "cells": [c("Team total")] + [c(f"=SUM({L}2:{L}6)") for L in "BCDEFG"] + [c()]})
    return {
        "id": "sales-tracker", "title": "Sales tracker", "category": "Sales",
        "description": "Quarterly sales by rep with totals, averages and a top performer flag.",
        "workbook": {"title": "Sales Tracker", "sheets": [
            {"name": "Sales", "columns": [col("Rep")] + [col(q, "currency") for q in ("Q1", "Q2", "Q3", "Q4")]
                                         + [col("Total", "currency"), col("Quarterly avg", "currency"), col("Flag")],
             "rows": rows,
             "charts": [{"type": "column", "title": "Total sales by rep", "categories": "A2:A6",
                         "series": [{"name": "Total", "range": "F2:F6"}]}]},
        ]},
    }


def projects():
    tasks = [("Kickoff and scope", "Ava", "2026-10-01", "2026-10-07", "Done"),
             ("Requirements", "Liam", "2026-10-06", "2026-10-17", "Done"),
             ("Design", "Sofia", "2026-10-15", "2026-11-07", "In progress"),
             ("Build", "Noah", "2026-11-01", "2026-12-12", "Not started"),
             ("Testing", "Mia", "2026-12-08", "2026-12-23", "Not started"),
             ("Launch", "Ava", "2027-01-05", "2027-01-08", "Not started")]
    rows = [{"cells": [c(t), c(o), c(s), c(e), c(st), c(f"=D{i + 2}-C{i + 2}"),
                       c(f'=IF(E{i + 2}="Done",0,MAX(D{i + 2}-TODAY(),0))')]} for i, (t, o, s, e, st) in enumerate(tasks)]
    return {
        "id": "project-timeline", "title": "Project timeline", "category": "Projects",
        "description": "Tasks, owners, dates, status dropdown and days remaining.",
        "workbook": {"title": "Project Timeline", "sheets": [
            {"name": "Timeline", "columns": [col("Task", width=24), col("Owner"), col("Start", "date"), col("End", "date"),
                                             col("Status"), col("Duration (days)", "integer"), col("Days remaining", "integer")],
             "rows": rows,
             "validations": [{"range": "E2:E200", "type": "list", "options": ["Not started", "In progress", "Blocked", "Done"]}]},
        ]},
    }


def main():
    OUT.mkdir(exist_ok=True)
    for make in (budget, invoice, loan, inventory, timesheet, sales, projects):
        t = make()
        WorkbookSpec.model_validate(t["workbook"])  # fail loudly on a typo
        (OUT / f"{t['id']}.json").write_text(json.dumps(t, indent=1))
        print("wrote", t["id"])


if __name__ == "__main__":
    main()
