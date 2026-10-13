"""Неделя 5: несколько фактов, active/inactive, USERELATIONSHIP, 1:1, M:M через bridge.

Эксперимент Single → Both (задание 17) делается в интерфейсе: «Изменить связь».
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from pbip import Query, build
from pbir import col, measure, proj, textbox, visual, write_pages

BASE = os.path.expanduser("~/PowerBI_2026/Week05/_pbip/DS24-B_Polienko_Week5")
MONEY = "#,0"
M_TYPES = {"text": "type text", "int": "Int64.Type", "num": "type number", "date": "type date", "pct": "Percentage.Type"}


def table(name, cols, **kw):
    types = ", ".join(f'{{"{c}", {M_TYPES[t]}}}' for c, t in cols)
    return Query(name, f'''
        let
            Источник = Excel.Workbook(File.Contents(pSourceFile), null, true),
            {name}Table_Table = Источник{{[Item="{name}Table",Kind="Table"]}}[Data],
            #"Измененный тип" = Table.TransformColumnTypes({name}Table_Table,{{{types}}})
        in
            #"Измененный тип"''', "Модель", [(c, "num" if t == "pct" else t) for c, t in cols], **kw)


T, I, N, D = "text", "int", "num", "date"
queries = [
    table("FactSales", [("SalesLineID", T), ("OrderID", I), ("OrderDate", D), ("ShipDate", D), ("ProductID", T),
                        ("CustomerID", T), ("StoreID", T), ("Quantity", I), ("UnitPrice", N), ("DiscountPct", "pct"),
                        ("SalesAmount", N), ("CostAmount", N), ("ProfitAmount", N)],
          hidden={"SalesLineID", "ProductID", "CustomerID", "StoreID"},
          fmt={"UnitPrice": MONEY, "SalesAmount": MONEY, "CostAmount": MONEY, "ProfitAmount": MONEY, "DiscountPct": "0%"},
          measures={
              "Total Sales": ("SUM(FactSales[SalesAmount])", MONEY),
              "Sales by Ship Date": ("CALCULATE([Total Sales], USERELATIONSHIP(DimDate[Date], FactSales[ShipDate]))", MONEY),
              "Ship - Order diff": ("[Sales by Ship Date] - [Total Sales]", MONEY),
              "Total Returns": ("SUM(FactReturns[ReturnAmount])", MONEY),
              "Return Rate %": ("DIVIDE([Total Returns], [Total Sales])", "0.0%"),
              "Total Budget": ("SUM(FactBudget[BudgetAmount])", MONEY),
              "Sales vs Budget %": ("DIVIDE([Total Sales], [Total Budget])", "0.0%"),
              "Promo Links": "COUNTROWS(BridgeProductPromotion)",
              "Customers (Single)": "COUNTROWS(DimCustomer)",
              "Customers (Both via CROSSFILTER)":
                  "CALCULATE(COUNTROWS(DimCustomer), CROSSFILTER(FactSales[CustomerID], DimCustomer[CustomerID], Both))",
          }),
    table("FactReturns", [("ReturnID", T), ("OrderID", I), ("ReturnDate", D), ("ProductID", T), ("CustomerID", T),
                          ("StoreID", T), ("QuantityReturned", I), ("ReturnAmount", N), ("Reason", T)],
          hidden={"ReturnID", "ProductID", "CustomerID", "StoreID"}, fmt={"ReturnAmount": MONEY}),
    table("FactBudget", [("BudgetID", T), ("MonthStart", D), ("StoreID", T), ("CategoryID", T), ("BudgetAmount", N)],
          hidden={"BudgetID", "StoreID", "CategoryID"}, fmt={"BudgetAmount": MONEY}),
    table("DimDate", [("DateKey", I), ("Date", D), ("Year", I), ("Quarter", T), ("MonthNo", I), ("MonthName", T),
                      ("ISOWeek", I), ("Day", I), ("IsWorkday", I), ("MonthStart", D)], hidden={"DateKey"}),
    table("DimCategory", [("CategoryID", T), ("CategoryName", T)], hidden={"CategoryID"}),
    table("DimProduct", [("ProductID", T), ("ProductName", T), ("CategoryID", T), ("CategoryName", T), ("Brand", T),
                         ("ListPrice", N), ("StandardCost", N)], hidden={"ProductID", "CategoryID"},
          fmt={"ListPrice": MONEY, "StandardCost": MONEY}),
    table("DimCustomer", [("CustomerID", T), ("CustomerName", T), ("Segment", T), ("CityID", T), ("City", T),
                          ("RegionID", T), ("Region", T)], hidden={"CustomerID"}),
    table("DimStore", [("StoreID", T), ("StoreName", T), ("CityID", T), ("City", T), ("RegionID", T), ("Region", T)],
          hidden={"StoreID"}),
    table("DimStoreDetails", [("StoreID", T), ("StoreFormat", T), ("AreaSqM", I), ("OpenDate", D), ("Ownership", T)],
          hidden={"StoreID"}),
    table("DimPromotion", [("PromotionID", T), ("PromotionName", T), ("StartDate", D), ("EndDate", D), ("PromotionType", T)],
          hidden={"PromotionID"}),
    table("BridgeProductPromotion", [("ProductID", T), ("PromotionID", T)], hidden={"ProductID", "PromotionID"}),
]

rel = lambda many, one, *opts: [many, one, *opts]
relationships = [
    rel("FactSales[ProductID]", "DimProduct[ProductID]"),
    rel("FactSales[CustomerID]", "DimCustomer[CustomerID]"),
    rel("FactSales[StoreID]", "DimStore[StoreID]"),
    rel("FactSales[OrderDate]", "DimDate[Date]"),
    rel("FactSales[ShipDate]", "DimDate[Date]", "isActive: false"),
    rel("FactReturns[ProductID]", "DimProduct[ProductID]"),
    rel("FactReturns[CustomerID]", "DimCustomer[CustomerID]"),
    rel("FactReturns[StoreID]", "DimStore[StoreID]"),
    rel("FactReturns[ReturnDate]", "DimDate[Date]"),
    rel("DimStoreDetails[StoreID]", "DimStore[StoreID]", "fromCardinality: one", "crossFilteringBehavior: bothDirections"),
    rel("BridgeProductPromotion[ProductID]", "DimProduct[ProductID]", "crossFilteringBehavior: bothDirections"),
    rel("BridgeProductPromotion[PromotionID]", "DimPromotion[PromotionID]"),
    rel("FactBudget[StoreID]", "DimStore[StoreID]"),
    rel("FactBudget[CategoryID]", "DimCategory[CategoryID]"),
    rel("FactBudget[MonthStart]", "DimDate[Date]"),
    rel("DimProduct[CategoryID]", "DimCategory[CategoryID]"),
]

if __name__ == "__main__":
    build(BASE, queries, relationships, params={"pSourceFile": r"Z:\Week05\PowerBI_Week5_Relationships_Dataset.xlsx"},
          sort_by={("DimDate", "MonthName"): "MonthNo"},
          layout={"DimDate": (620, 0), "FactSales": (620, 380), "FactReturns": (620, 760), "FactBudget": (1300, 0),
                  "DimProduct": (1000, 380), "DimCategory": (1300, 380), "BridgeProductPromotion": (1000, 760),
                  "DimPromotion": (1300, 760), "DimCustomer": (240, 380), "DimStore": (240, 760), "DimStoreDetails": (-140, 760)})

    m = lambda n: proj(measure("FactSales", n))
    c = lambda t, f: proj(col(t, f))
    desc = lambda n: (measure("FactSales", n), "Descending")
    write_pages(BASE + ".Report", [
        ("Shared dimensions", [
            textbox((20, 10, 1880, 80), "Продажи и возвраты через общие измерения (FactSales + FactReturns)", 24),
            visual("card", (20, 100, 440, 160), {"Values": [m("Total Sales")]}, "Total Sales"),
            visual("card", (480, 100, 440, 160), {"Values": [m("Total Returns")]}, "Total Returns"),
            visual("card", (940, 100, 440, 160), {"Values": [m("Return Rate %")]}, "Return Rate %"),
            visual("slicer", (1400, 100, 500, 400), {"Values": [c("DimProduct", "CategoryName")]}, "DimProduct: CategoryName"),
            visual("tableEx", (20, 280, 1360, 780), {"Values": [c("DimProduct", "ProductName"), c("DimProduct", "Brand"),
                                                               m("Total Sales"), m("Total Returns"), m("Return Rate %")]},
                   "ProductName + Total Sales + Total Returns", desc("Total Sales")),
            visual("clusteredBarChart", (1400, 520, 500, 540), {"Category": [c("DimStore", "StoreName")],
                                                                "Y": [m("Total Sales"), m("Total Returns")]},
                   "DimStore фильтрует оба факта", desc("Total Sales")),
        ]),
        ("OrderDate vs ShipDate", [
            textbox((20, 10, 1880, 80), "Активная связь OrderDate и неактивная ShipDate (USERELATIONSHIP)", 24),
            visual("slicer", (20, 100, 300, 300), {"Values": [c("DimDate", "Year")]}, "DimDate: Year"),
            visual("lineChart", (340, 100, 1560, 560), {"Category": [c("DimDate", "MonthStart")],
                                                        "Y": [m("Total Sales"), m("Sales by Ship Date")]},
                   "Total Sales (по дате заказа) и Sales by Ship Date (по дате отгрузки)"),
            visual("tableEx", (20, 680, 1880, 380), {"Values": [c("DimDate", "MonthStart"), m("Total Sales"),
                                                               m("Sales by Ship Date"), m("Ship - Order diff")]},
                   "Разница по месяцам", (col("DimDate", "MonthStart"), "Ascending")),
        ]),
        ("1:1 и Many-to-Many", [
            textbox((20, 10, 1880, 80), "DimStore ↔ DimStoreDetails (1:1) и Product–Promotion через bridge", 24),
            visual("tableEx", (20, 100, 900, 460), {"Values": [c("DimStore", "StoreName"), c("DimStoreDetails", "StoreFormat"),
                                                              c("DimStoreDetails", "AreaSqM"), m("Total Sales")]},
                   "StoreName + StoreFormat + AreaSqM (1:1)", desc("Total Sales")),
            visual("slicer", (940, 100, 400, 460), {"Values": [c("DimPromotion", "PromotionName")]}, "DimPromotion: PromotionName"),
            visual("clusteredBarChart", (1360, 100, 540, 460), {"Category": [c("DimPromotion", "PromotionName")], "Y": [m("Promo Links")]},
                   "Товаров в акции", desc("Promo Links")),
            visual("tableEx", (20, 580, 1880, 480), {"Values": [c("DimPromotion", "PromotionName"), c("DimProduct", "ProductName"),
                                                               c("DimProduct", "Brand"), m("Promo Links"), m("Total Sales")]},
                   "PromotionName × ProductName (bridge)", (col("DimPromotion", "PromotionName"), "Ascending")),
        ]),
        ("Single vs Both", [
            textbox((20, 10, 1880, 80), "Направление фильтра: срез по товару и число клиентов", 24),
            visual("slicer", (20, 100, 500, 500), {"Values": [c("DimProduct", "ProductName")]}, "DimProduct: ProductName"),
            visual("card", (540, 100, 660, 220), {"Values": [m("Customers (Single)")]}, "COUNTROWS(DimCustomer)"),
            visual("card", (1220, 100, 680, 220), {"Values": [m("Customers (Both via CROSSFILTER)")]}, "То же с CROSSFILTER(..., Both)"),
            visual("tableEx", (540, 340, 1360, 720), {"Values": [c("DimCustomer", "Region"), m("Customers (Single)"),
                                                                m("Customers (Both via CROSSFILTER)"), m("Total Sales")]},
                   "По регионам клиентов"),
        ]),
        ("Sales vs Budget", [
            textbox((20, 10, 1880, 80), "FactBudget (магазин × категория × месяц) против FactSales", 24),
            visual("slicer", (20, 100, 300, 300), {"Values": [c("DimDate", "Year")]}, "DimDate: Year"),
            visual("card", (340, 100, 500, 160), {"Values": [m("Total Sales")]}, "Total Sales"),
            visual("card", (860, 100, 500, 160), {"Values": [m("Total Budget")]}, "Total Budget"),
            visual("card", (1380, 100, 520, 160), {"Values": [m("Sales vs Budget %")]}, "Sales vs Budget %"),
            visual("clusteredColumnChart", (340, 280, 1560, 380), {"Category": [c("DimCategory", "CategoryName")],
                                                                   "Y": [m("Total Sales"), m("Total Budget")]},
                   "По категориям (DimCategory)", desc("Total Budget")),
            visual("clusteredColumnChart", (20, 680, 1880, 380), {"Category": [c("DimStore", "StoreName")],
                                                                  "Y": [m("Total Sales"), m("Total Budget")]},
                   "По магазинам (DimStore)", desc("Total Budget")),
        ]),
    ])
    print("ok")
