"""Неделя 6: DAX — таблица _Measures, базовые KPI, итераторы, VAR/RETURN, IF/SWITCH, branching."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from pbip import Query, build
from pbir import col, measure, proj, textbox, visual, write_pages
import week05 as w5  # та же схема таблиц, что и на неделе 5

BASE = os.path.expanduser("~/PowerBI_2026/Week06/_pbip/DS24-B_Polienko_Week6")
MONEY, PCT, INT = "#,0", "0.0%", "#,0"
B, K, IT, L, R = "1. Базовые", "2. KPI", "3. Итераторы", "4. Логика", "5. Возвраты и динамика"

# из недели 5 берём только нужные по заданию таблицы, меры — в _Measures
keep = ["FactSales", "FactReturns", "DimProduct", "DimCustomer", "DimStore", "DimDate"]
tables = {q.name: q for q in w5.queries if q.name in keep}
for t in tables.values():
    t.measures = {}
fact = tables["FactSales"]
fact.calc = {"Profit Row Calc": ("FactSales[SalesAmount] - FactSales[CostAmount]", "num", MONEY)}
fact.hidden.add("Profit Row Calc")  # задание 17: временный столбец, после сравнения скрыт

measures = Query("_Measures", '''
    let
        Источник = #table(type table [Служебный = Int64.Type], {{1}})
    in
        Источник''', None, [("Служебный", "int")], hidden={"Служебный"}, measures={
    # 1. Базовый слой — только здесь SUM/COUNT по столбцам
    "Total Sales": ("SUM(FactSales[SalesAmount])", MONEY, B),
    "Total Profit": ("SUM(FactSales[ProfitAmount])", MONEY, B),
    "Total Cost": ("SUM(FactSales[CostAmount])", MONEY, B),
    "Units Sold": ("SUM(FactSales[Quantity])", INT, B),
    "Orders Count": ("DISTINCTCOUNT(FactSales[OrderID])", INT, B),
    "Sales Lines": ("COUNTROWS(FactSales)", INT, B),
    "Customers Count": ("DISTINCTCOUNT(FactSales[CustomerID])", INT, B),
    "Total Returns": ("SUM(FactReturns[ReturnAmount])", MONEY, B),
    # 2. Производные KPI — только через базовые меры
    "Average Sale Row": ("AVERAGE(FactSales[SalesAmount])", MONEY, K),
    "Average Order Value": ("DIVIDE([Total Sales], [Orders Count], 0)", MONEY, K),
    "Profit Margin %": ("""
        VAR SalesValue = [Total Sales]
        VAR ProfitValue = [Total Profit]
        RETURN
            DIVIDE(ProfitValue, SalesValue, 0)""", PCT, K),
    "Sales per Customer": ("DIVIDE([Total Sales], [Customers Count], 0)", MONEY, K),
    "Units per Order": ("DIVIDE([Units Sold], [Orders Count], 0)", "0.00", K),
    # 3. Итераторы (Row Context внутри)
    "Sales Recalculated": ("""
        SUMX(
            FactSales,
            FactSales[Quantity] * FactSales[UnitPrice] * (1 - FactSales[DiscountPct])
        )""", MONEY, IT),
    "Average Revenue per Line": ("""
        AVERAGEX(
            FactSales,
            FactSales[Quantity] * FactSales[UnitPrice] * (1 - FactSales[DiscountPct])
        )""", MONEY, IT),
    "Discount Amount": ("""
        SUMX(
            FactSales,
            FactSales[Quantity] * FactSales[UnitPrice] * FactSales[DiscountPct]
        )""", MONEY, IT),
    "Discount Rate %": ("DIVIDE([Discount Amount], [Total Sales] + [Discount Amount], 0)", PCT, IT),
    # 4. Логика
    "Sales Status": ('IF([Total Sales] >= 10000000, "Высокие продажи", "Обычные продажи")', None, L),
    "Performance Level": ("""
        SWITCH(
            TRUE(),
            [Profit Margin %] >= 0.25, "A",
            [Profit Margin %] >= 0.15, "B",
            [Profit Margin %] >= 0.05, "C",
            "D"
        )""", None, L),
    "Return Alert": ('IF([Return Rate %] > 0.08, "Проверить возвраты", "Норма")', None, L),
    "AOV Segment": ("""
        SWITCH(
            TRUE(),
            ISBLANK([Orders Count]), BLANK(),
            [Average Order Value] >= 1000000, "Крупный чек",
            [Average Order Value] >= 500000, "Средний чек",
            "Малый чек"
        )""", None, L),
    # 5. Возвраты и динамика
    "Return Rate %": ("DIVIDE([Total Returns], [Total Sales], 0)", PCT, R),
    "Sales PY": ("CALCULATE([Total Sales], SAMEPERIODLASTYEAR(DimDate[Date]))", MONEY, R),
    "Sales YoY %": ("""
        VAR CurrentSales = [Total Sales]
        VAR PreviousSales = [Sales PY]
        VAR Growth = CurrentSales - PreviousSales
        RETURN
            IF(HASONEVALUE(DimDate[Year]) && NOT ISBLANK(PreviousSales), DIVIDE(Growth, PreviousSales))""", PCT, R),
})

queries = [measures] + [tables[n] for n in keep]
rels = [r for r in w5.relationships if r[0].split("[")[0] in keep and r[1].split("[")[0] in keep]
build(BASE, queries, rels, params={"pSourceFile": r"Z:\Week06\PowerBI_Week6_DAX_Dataset.xlsx"},
      sort_by={("DimDate", "MonthName"): "MonthNo"},
      layout={"_Measures": (0, 0), "DimDate": (620, 0), "FactSales": (620, 380), "FactReturns": (620, 760),
              "DimProduct": (1000, 380), "DimCustomer": (240, 380), "DimStore": (240, 760)})

m = lambda n: proj(measure("_Measures", n))
c = lambda t, f: proj(col(t, f))
desc = lambda n: (measure("_Measures", n), "Descending")
cards = lambda names, y=100, h=150: [visual("card", (20 + i * (1880 // len(names)), y, 1880 // len(names) - 20, h),
                                            {"Values": [m(n)]}, n) for i, n in enumerate(names)]
write_pages(BASE + ".Report", [
    ("Базовые KPI", [
        textbox((20, 10, 1880, 80), "Неделя 6 — базовые меры и KPI (таблица _Measures)", 24),
        *cards(["Total Sales", "Total Profit", "Units Sold", "Orders Count", "Sales Lines", "Average Order Value", "Profit Margin %"]),
        visual("tableEx", (20, 270, 1000, 380), {"Values": [c("DimStore", "StoreName"), m("Average Sale Row"),
                                                           m("Average Order Value"), m("Total Sales"), m("Sales Status")]},
               "По магазинам: Average Sale Row vs Average Order Value", desc("Total Sales")),
        visual("pivotTable", (1040, 270, 860, 380), {"Rows": [c("DimProduct", "CategoryName")], "Columns": [c("DimStore", "Region")],
                                                     "Values": [m("Profit Margin %")]}, "Profit Margin %: категория × регион магазина"),
        visual("tableEx", (20, 670, 1000, 390), {"Values": [c("DimProduct", "ProductName"), c("DimProduct", "Brand"),
                                                           m("Total Sales"), m("Sales Status")]},
               "IF: Sales Status (порог 10 000 000) по товарам", desc("Total Sales")),
        visual("tableEx", (1040, 670, 860, 390), {"Values": [m("Orders Count"), m("Sales Lines"), m("Customers Count"),
                                                            m("Units per Order"), m("Sales per Customer")]},
               "DISTINCTCOUNT vs COUNTROWS"),
    ]),
    ("Filter Context", [
        textbox((20, 10, 1880, 80), "Одна мера [Total Sales] — разные Filter Context", 24),
        visual("slicer", (20, 100, 300, 220), {"Values": [c("DimDate", "Year")]}, "Year"),
        visual("slicer", (20, 340, 300, 340), {"Values": [c("DimProduct", "Brand")]}, "Brand"),
        visual("slicer", (20, 700, 300, 360), {"Values": [c("DimCustomer", "Segment")]}, "Segment"),
        visual("card", (340, 100, 500, 220), {"Values": [m("Total Sales")]}, "Total Sales (карточка)"),
        visual("card", (860, 100, 500, 220), {"Values": [m("Orders Count")]}, "Orders Count"),
        visual("card", (1380, 100, 520, 220), {"Values": [m("Profit Margin %")]}, "Profit Margin %"),
        visual("tableEx", (340, 340, 620, 720), {"Values": [c("DimProduct", "CategoryName"), m("Total Sales")]},
               "CategoryName | Total Sales", desc("Total Sales")),
        visual("pivotTable", (980, 340, 920, 720), {"Rows": [c("DimCustomer", "Region")], "Columns": [c("DimCustomer", "Segment")],
                                                    "Values": [m("Total Sales")]}, "Region × Segment | Total Sales"),
    ]),
    ("Итераторы и логика", [
        textbox((20, 10, 1880, 80), "SUMX / AVERAGEX, VAR, IF, SWITCH", 24),
        visual("tableEx", (20, 100, 1880, 330), {"Values": [c("DimProduct", "CategoryName"), m("Total Sales"), m("Sales Recalculated"),
                                                           m("Average Sale Row"), m("Average Revenue per Line"), m("Discount Amount"),
                                                           m("Profit Margin %"), m("Performance Level")]},
               "По категориям: итераторы и SWITCH Performance Level", desc("Total Sales")),
        visual("tableEx", (20, 450, 930, 300), {"Values": [c("DimProduct", "CategoryName"), m("Average Order Value"), m("AOV Segment"),
                                                          m("Return Rate %"), m("Return Alert")]},
               "Свои IF/SWITCH по категориям", desc("Average Order Value")),
        visual("tableEx", (970, 450, 930, 300), {"Values": [c("DimProduct", "Brand"), m("Average Order Value"), m("AOV Segment")]},
               "AOV Segment по брендам", desc("Average Order Value")),
        visual("tableEx", (20, 770, 930, 290), {"Values": [c("DimStore", "StoreName"), m("Total Returns"), m("Return Rate %"),
                                                          m("Return Alert")]},
               "Return Alert (порог 8%) по магазинам", desc("Return Rate %")),
        visual("tableEx", (970, 770, 930, 290), {"Values": [c("DimDate", "Year"), m("Total Sales"), m("Sales PY"), m("Sales YoY %")]},
               "VAR ×3: Sales YoY %"),
    ]),
    ("Measure vs Column", [
        textbox((20, 10, 1880, 80), "Row Context на строках FactSales и Measure против Calculated Column", 24),
        visual("tableEx", (20, 100, 1200, 420), {"Values": [c("FactSales", "SalesLineID"), c("FactSales", "Quantity"),
                                                           c("FactSales", "UnitPrice"), c("FactSales", "DiscountPct"),
                                                           c("FactSales", "SalesAmount"), m("Sales Recalculated")]},
               "Row Context: Quantity × UnitPrice × (1 − DiscountPct) по строкам", (col("FactSales", "SalesLineID"), "Ascending")),
        visual("slicer", (1240, 100, 660, 420), {"Values": [c("DimProduct", "CategoryName")]}, "CategoryName"),
        visual("tableEx", (20, 540, 1880, 520), {"Values": [c("DimStore", "StoreName"), m("Total Profit"),
                                                           proj({"Aggregation": {"Expression": col("FactSales", "Profit Row Calc"), "Function": 0}},
                                                                "Sum of Profit Row Calc (столбец)"),
                                                           m("Profit Margin %")]},
               "Measure [Total Profit] и SUM вычисляемого столбца Profit Row Calc", desc("Total Profit")),
    ]),
])
print("ok")
