"""Неделя 4: Flat Table → Star Schema (FactSales + 5 измерений)."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from pbip import Query, build
from pbir import col, measure, proj, textbox, visual, write_pages

BASE = os.path.expanduser("~/PowerBI_2026/Week04/_pbip/DS24-B_Polienko_Week4")
MONEY = "#,0"

stg = Query("stg_Sales_Flat", '''
    let
        Источник = Excel.Workbook(File.Contents(pSourceFile), null, true),
        Sales_Flat_Sheet = Источник{[Item="Sales_Flat",Kind="Sheet"]}[Data],
        #"Повышенные заголовки" = Table.PromoteHeaders(Sales_Flat_Sheet, [PromoteAllScalars=true]),
        #"Измененный тип" = Table.TransformColumnTypes(#"Повышенные заголовки",{{"SalesLineID", type text}, {"OrderID", Int64.Type}, {"OrderDate", type date}, {"ProductID", type text}, {"ProductName", type text}, {"CategoryID", type text}, {"CategoryName", type text}, {"Brand", type text}, {"CustomerID", type text}, {"CustomerName", type text}, {"CustomerSegment", type text}, {"CustomerCityID", type text}, {"CustomerCity", type text}, {"CustomerRegionID", type text}, {"CustomerRegion", type text}, {"StoreID", type text}, {"StoreName", type text}, {"StoreCityID", type text}, {"StoreCity", type text}, {"StoreRegionID", type text}, {"StoreRegion", type text}, {"ManagerID", type text}, {"ManagerName", type text}, {"Quantity", Int64.Type}, {"UnitPrice", Currency.Type}, {"DiscountPct", Percentage.Type}, {"SalesAmount", Currency.Type}, {"CostAmount", Currency.Type}, {"ProfitAmount", Currency.Type}})
    in
        #"Измененный тип"''', "00_STAGING", load=False)


def dim(name, key, cols, extra_steps=""):
    keep = ", ".join(f'"{c}"' for c, _ in cols if c in ALL)
    last = '#"Удаленные дубликаты"'
    return Query(name, f'''
        let
            Источник = stg_Sales_Flat,
            #"Другие удаленные столбцы" = Table.SelectColumns(Источник,{{{keep}}}),
            #"Удаленные дубликаты" = Table.Distinct(#"Другие удаленные столбцы", {{"{key}"}}){"," if extra_steps else ""}{extra_steps}
        in
            {extra_steps.strip().rsplit(chr(10), 1)[-1].split(" = ")[0].strip() if extra_steps else last}''',
        "10_MODEL", cols, hidden={key} - {"OrderDate"})


ALL = {"ProductID", "ProductName", "CategoryID", "CategoryName", "Brand", "CustomerID", "CustomerName", "CustomerSegment",
       "CustomerCityID", "CustomerCity", "CustomerRegionID", "CustomerRegion", "StoreID", "StoreName", "StoreCityID",
       "StoreCity", "StoreRegionID", "StoreRegion", "ManagerID", "ManagerName", "OrderDate"}
T = "text"

DATE_STEPS = '''
            #"Вставленный год" = Table.AddColumn(#"Удаленные дубликаты", "Year", each Date.Year([OrderDate]), Int64.Type),
            #"Вставленный месяц" = Table.AddColumn(#"Вставленный год", "MonthNo", each Date.Month([OrderDate]), Int64.Type),
            #"Вставленное название месяца" = Table.AddColumn(#"Вставленный месяц", "MonthName", each Date.MonthName([OrderDate], "ru-RU"), type text),
            #"Вставленный квартал" = Table.AddColumn(#"Вставленное название месяца", "Quarter", each Date.QuarterOfYear([OrderDate]), Int64.Type),
            #"Сортированные строки" = Table.Sort(#"Вставленный квартал",{{"OrderDate", Order.Ascending}})'''

queries = [
    stg,
    dim("DimProduct", "ProductID", [("ProductID", T), ("ProductName", T), ("CategoryID", T), ("CategoryName", T), ("Brand", T)]),
    dim("DimCustomer", "CustomerID", [("CustomerID", T), ("CustomerName", T), ("CustomerSegment", T), ("CustomerCityID", T),
                                      ("CustomerCity", T), ("CustomerRegionID", T), ("CustomerRegion", T)]),
    dim("DimStore", "StoreID", [("StoreID", T), ("StoreName", T), ("StoreCityID", T), ("StoreCity", T), ("StoreRegionID", T),
                                ("StoreRegion", T)]),
    dim("DimManager", "ManagerID", [("ManagerID", T), ("ManagerName", T)]),
    dim("DimDate", "OrderDate", [("OrderDate", "date"), ("Year", "int"), ("MonthNo", "int"), ("MonthName", T), ("Quarter", "int")],
        DATE_STEPS),
    Query("FactSales", '''
        let
            Источник = stg_Sales_Flat,
            #"Другие удаленные столбцы" = Table.SelectColumns(Источник,{"SalesLineID", "OrderID", "OrderDate", "ProductID", "CustomerID", "StoreID", "ManagerID", "Quantity", "UnitPrice", "DiscountPct", "SalesAmount", "CostAmount", "ProfitAmount"})
        in
            #"Другие удаленные столбцы"''', "10_MODEL",
          [("SalesLineID", T), ("OrderID", "int"), ("OrderDate", "date"), ("ProductID", T), ("CustomerID", T), ("StoreID", T),
           ("ManagerID", T), ("Quantity", "int"), ("UnitPrice", "num"), ("DiscountPct", "num"), ("SalesAmount", "num"),
           ("CostAmount", "num"), ("ProfitAmount", "num")],
          hidden={"SalesLineID", "ProductID", "CustomerID", "StoreID", "ManagerID"},
          fmt={"UnitPrice": MONEY, "SalesAmount": MONEY, "CostAmount": MONEY, "ProfitAmount": MONEY, "DiscountPct": "0%"},
          measures={
              "Total Sales": ("SUM(FactSales[SalesAmount])", MONEY),
              "Total Profit": ("SUM(FactSales[ProfitAmount])", MONEY),
              "Sales Lines": "COUNTROWS(FactSales)",
              "Orders": "DISTINCTCOUNT(FactSales[OrderID])",
              # аудит PK: строк = distinct ключей
              "PK DimProduct rows | distinct": 'COUNTROWS(DimProduct) & " | " & DISTINCTCOUNT(DimProduct[ProductID])',
              "PK DimCustomer rows | distinct": 'COUNTROWS(DimCustomer) & " | " & DISTINCTCOUNT(DimCustomer[CustomerID])',
              "PK DimStore rows | distinct": 'COUNTROWS(DimStore) & " | " & DISTINCTCOUNT(DimStore[StoreID])',
              "PK DimManager rows | distinct": 'COUNTROWS(DimManager) & " | " & DISTINCTCOUNT(DimManager[ManagerID])',
              "PK DimDate rows | distinct": 'COUNTROWS(DimDate) & " | " & DISTINCTCOUNT(DimDate[OrderDate])',
              # аудит FK: строк факта без пары в измерении
              "FK missing Product": "COUNTROWS(FILTER(FactSales, ISBLANK(RELATED(DimProduct[ProductID])))) + 0",
              "FK missing Customer": "COUNTROWS(FILTER(FactSales, ISBLANK(RELATED(DimCustomer[CustomerID])))) + 0",
              "FK missing Store": "COUNTROWS(FILTER(FactSales, ISBLANK(RELATED(DimStore[StoreID])))) + 0",
              "FK missing Manager": "COUNTROWS(FILTER(FactSales, ISBLANK(RELATED(DimManager[ManagerID])))) + 0",
              "FK missing Date": "COUNTROWS(FILTER(FactSales, ISBLANK(RELATED(DimDate[OrderDate])))) + 0",
          }),
]

build(BASE, queries,
      relationships=[f"FactSales[{k}] {d}[{k}]".split() for k, d in
                     [("ProductID", "DimProduct"), ("CustomerID", "DimCustomer"), ("StoreID", "DimStore"),
                      ("ManagerID", "DimManager"), ("OrderDate", "DimDate")]],
      params={"pSourceFile": r"Z:\Week04\PowerBI_Week4_Flat_Sales_Student.xlsx"},
      sort_by={("DimDate", "MonthName"): "MonthNo"},
      layout={"FactSales": (520, 330), "DimDate": (520, 0), "DimProduct": (900, 330), "DimCustomer": (140, 330),
              "DimStore": (330, 700), "DimManager": (710, 700)})

m = lambda n: proj(measure("FactSales", n))
write_pages(BASE + ".Report", [
    ("Проверка модели", [
        textbox((20, 10, 1880, 80), "Неделя 4 — Star Schema: тест фильтрации через измерения", 26),
        visual("card", (20, 100, 440, 160), {"Values": [m("Total Sales")]}, "Total Sales"),
        visual("card", (480, 100, 440, 160), {"Values": [m("Total Profit")]}, "Total Profit"),
        visual("card", (940, 100, 440, 160), {"Values": [m("Sales Lines")]}, "Строк факта (SalesLineID)"),
        visual("card", (1400, 100, 500, 160), {"Values": [m("Orders")]}, "Заказов (DISTINCTCOUNT OrderID)"),
        visual("slicer", (20, 280, 440, 250), {"Values": [proj(col("DimProduct", "CategoryName"))]}, "DimProduct: CategoryName"),
        visual("slicer", (20, 550, 440, 250), {"Values": [proj(col("DimCustomer", "CustomerRegion"))]}, "DimCustomer: CustomerRegion"),
        visual("slicer", (20, 820, 440, 240), {"Values": [proj(col("DimStore", "StoreRegion"))]}, "DimStore: StoreRegion"),
        visual("tableEx", (480, 280, 700, 780), {"Values": [proj(col("DimProduct", "ProductName")), proj(col("DimProduct", "Brand")),
                                                          m("Total Sales")]}, "ProductName + SalesAmount",
               (measure("FactSales", "Total Sales"), "Descending")),
        visual("clusteredColumnChart", (1200, 280, 700, 380), {"Category": [proj(col("DimDate", "Year"))], "Y": [m("Total Sales")]},
               "Sales by Year (DimDate)"),
        visual("clusteredBarChart", (1200, 680, 700, 380), {"Category": [proj(col("DimManager", "ManagerName"))], "Y": [m("Total Sales")]},
               "Sales by Manager (DimManager)", (measure("FactSales", "Total Sales"), "Descending")),
    ]),
    ("Аудит ключей", [
        textbox((20, 10, 1880, 80), "Аудит PK (строк | distinct) и FK (строк факта без пары в измерении)", 24),
        visual("multiRowCard", (20, 100, 920, 900), {"Values": [m(n) for n in (
            "PK DimProduct rows | distinct", "PK DimCustomer rows | distinct", "PK DimStore rows | distinct",
            "PK DimManager rows | distinct", "PK DimDate rows | distinct")]}, "Primary Keys"),
        visual("multiRowCard", (960, 100, 940, 900), {"Values": [m(n) for n in (
            "FK missing Product", "FK missing Customer", "FK missing Store", "FK missing Manager", "FK missing Date")]},
               "Foreign Keys без пары"),
    ]),
])
print("ok")
