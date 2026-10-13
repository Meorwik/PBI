"""Неделя 2: Power Query ETL над Week2_Dirty_Sales_Student.xlsx.

python3 week02.py        — итоговый проект
python3 week02.py demo   — плюс временные запросы для скриншотов (Fill Up, Duplicate)
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from pbip import Query, build
from pbir import agg, col, measure, proj, textbox, visual, write_pages, lit

BASE = os.path.expanduser("~/PowerBI_2026/Week02/_pbip/DS24-B_Polienko_Week2")
SRC = r"Z:\Week02\Week2_Dirty_Sales_Student.xlsx"
DEMO = "demo" in sys.argv[1:]

BLANK = 'each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))'


def sheet(name):
    return f'''
        Источник = src_Excel,
        {name}_Sheet = Источник{{[Item="{name}",Kind="Sheet"]}}[Data],
        #"Повышенные заголовки" = Table.PromoteHeaders({name}_Sheet, [PromoteAllScalars=true]),
        #"Удаленные пустые строки" = Table.SelectRows(#"Повышенные заголовки", {BLANK}),'''


STG_SALES_STEPS = sheet("Sales_Raw") + '''
        #"Тип текстовых полей" = Table.TransformColumnTypes(#"Удаленные пустые строки",{{"OrderID", Int64.Type}, {"CustomerID", type text}, {"CustomerName", type text}, {"City", type text}, {"Region", type text}, {"ProductCode", type text}, {"ProductName", type text}, {"Category", type text}, {"PaymentMethod", type text}, {"Manager", type text}, {"Status", type text}}),
        #"Обрезанный текст" = Table.TransformColumns(#"Тип текстовых полей",{{"CustomerName", Text.Trim, type text}, {"City", Text.Trim, type text}, {"Region", Text.Trim, type text}, {"PaymentMethod", Text.Trim, type text}, {"Manager", Text.Trim, type text}, {"Status", Text.Trim, type text}}),
        #"Очищенный текст" = Table.TransformColumns(#"Обрезанный текст",{{"CustomerName", Text.Clean, type text}, {"City", Text.Clean, type text}}),
        #"Двойные пробелы в CustomerName" = Table.ReplaceValue(#"Очищенный текст","  "," ",Replacer.ReplaceText,{"CustomerName"}),
        #"Замененное значение N/A" = Table.ReplaceValue(#"Двойные пробелы в CustomerName","N/A",null,Replacer.ReplaceValue,{"OrderDate", "UnitPrice", "Discount", "SalesAmount", "Status"}),
        #"Замененное значение -" = Table.ReplaceValue(#"Замененное значение N/A","-",null,Replacer.ReplaceValue,{"UnitPrice", "Discount", "SalesAmount"}),
        #"Замененное значение error" = Table.ReplaceValue(#"Замененное значение -","error",null,Replacer.ReplaceValue,{"SalesAmount"}),
        #"Замененное значение wrong" = Table.ReplaceValue(#"Замененное значение error","wrong",null,Replacer.ReplaceValue,{"OrderDate"}),
        #"Каждое слово с прописной буквы" = Table.TransformColumns(#"Замененное значение wrong",{{"CustomerName", Text.Proper, type text}, {"City", Text.Proper, type text}, {"Region", Text.Proper, type text}, {"Status", Text.Proper, type text}}),
        #"Единый PaymentMethod" = Table.TransformColumns(#"Каждое слово с прописной буквы",{{"PaymentMethod", each let u = Text.Upper(_) in if u = "CARD" or u = "BANK CARD" then "Card" else if u = "CASH" then "Cash" else if u = "KASPI QR" then "Kaspi QR" else if u = "QR" then "QR" else _, type text}}),
        #"Discount из текста" = Table.TransformColumns(#"Единый PaymentMethod",{{"Discount", each if _ is text then Number.From(Text.Remove(_, {" ", "%"})) / 100 else _}}),
        #"SalesAmount из текста" = Table.TransformColumns(#"Discount из текста",{{"SalesAmount", each if _ is text then Number.From(Text.Trim(_)) else _}}),
        #"OrderDate из разных форматов" = Table.TransformColumns(#"SalesAmount из текста",{{"OrderDate", each if _ = null then null else if _ is date then _ else if _ is datetime then DateTime.Date(_) else try Date.FromText(_, [Format="dd.MM.yyyy"]) otherwise try Date.FromText(_, [Format="yyyy-MM-dd"]) otherwise try Date.FromText(_, [Format="dd/MM/yyyy"]) otherwise null, type date}}),
        #"Измененный тип" = Table.TransformColumnTypes(#"OrderDate из разных форматов",{{"Quantity", Int64.Type}, {"UnitPrice", type number}, {"Discount", type number}, {"SalesAmount", type number}}),
        #"Замененные ошибки" = Table.ReplaceErrorValues(#"Измененный тип", {{"UnitPrice", null}})'''

stg_sales = Query("stg_Sales", "\n    let" + STG_SALES_STEPS + ''',
        #"Удаленные дубликаты" = Table.Distinct(#"Замененные ошибки")
    in
        #"Удаленные дубликаты"''', "00_STAGING", load=False)

stg_regions = Query("stg_Regions", '\n    let' + sheet("Regions") + '''
        #"Измененный тип" = Table.TransformColumnTypes(#"Удаленные пустые строки",{{"Region", type text}, {"City", type text}, {"RegionalManager", type text}}),
        #"Заполнено вниз" = Table.FillDown(#"Измененный тип",{"Region", "RegionalManager"})
    in
        #"Заполнено вниз"''', "00_STAGING", load=False)

stg_products = Query("stg_Products", '\n    let' + sheet("Products") + '''
        #"Измененный тип" = Table.TransformColumnTypes(#"Удаленные пустые строки",{{"ProductCode", type text}, {"ProductName", type text}, {"Category", type text}, {"StandardPrice", Int64.Type}, {"SupplierCode", type text}}),
        #"Дублированный столбец" = Table.DuplicateColumn(#"Измененный тип", "ProductCode", "ProductCode - копия"),
        #"Разделить столбец по разделителю" = Table.SplitColumn(#"Дублированный столбец", "ProductCode - копия", Splitter.SplitTextByDelimiter("-", QuoteStyle.Csv), {"DepartmentCode", "ProductTypeCode", "ProductNumber"}),
        #"Измененный тип1" = Table.TransformColumnTypes(#"Разделить столбец по разделителю",{{"DepartmentCode", type text}, {"ProductTypeCode", type text}, {"ProductNumber", type text}}),
        #"Объединенные столбцы" = Table.CombineColumns(#"Измененный тип1",{"DepartmentCode", "ProductTypeCode"},Combiner.CombineTextByDelimiter("-", QuoteStyle.None),"ProductGroupCode"),
        #"Вставленный текст между разделителями" = Table.AddColumn(#"Объединенные столбцы", "ProductTypeFromExample", each Text.BetweenDelimiters([ProductCode], "-", "-"), type text)
    in
        #"Вставленный текст между разделителями"''', "00_STAGING", load=False)

stg_monthly = Query("stg_MonthlySales", '\n    let' + sheet("Monthly_Sales") + '''
        #"Измененный тип" = Table.TransformColumnTypes(#"Удаленные пустые строки",{{"Product", type text}, {"Jan", Int64.Type}, {"Feb", Int64.Type}, {"Mar", Int64.Type}, {"Apr", Int64.Type}, {"May", Int64.Type}, {"Jun", Int64.Type}, {"Jul", Int64.Type}, {"Aug", Int64.Type}})
    in
        #"Измененный тип"''', "00_STAGING", load=False)

src = Query("src_Excel", '''
    let
        Источник = Excel.Workbook(File.Contents(pSourceFile), null, true)
    in
        Источник''', "00_STAGING", load=False)

SALES_COLS = [("OrderID", "int"), ("OrderDate", "date"), ("CustomerID", "text"), ("CustomerName", "text"), ("City", "text"),
              ("Region", "text"), ("ProductCode", "text"), ("ProductName", "text"), ("Category", "text"), ("Quantity", "int"),
              ("UnitPrice", "num"), ("Discount", "num"), ("SalesAmount", "num"), ("PaymentMethod", "text"), ("Manager", "text"),
              ("Status", "text")]

fact = Query("FactSales", '''
    let
        Источник = stg_Sales,
        #"Объединенные запросы" = Table.NestedJoin(Источник, {"ProductCode"}, stg_Products, {"ProductCode"}, "stg_Products", JoinKind.LeftOuter),
        #"Развернутый элемент stg_Products" = Table.ExpandTableColumn(#"Объединенные запросы", "stg_Products", {"StandardPrice"}, {"StandardPrice"}),
        #"UnitPrice из прайса" = Table.ReplaceValue(#"Развернутый элемент stg_Products", each [UnitPrice], each if [UnitPrice] = null then [StandardPrice] else [UnitPrice], Replacer.ReplaceValue, {"UnitPrice"}),
        #"Восстановлен Discount" = Table.ReplaceValue(#"UnitPrice из прайса", each [Discount], each if [Discount] = null and [Quantity] <> null and [Quantity] > 0 and [SalesAmount] <> null then Number.Round(1 - [SalesAmount] / ([Quantity] * [UnitPrice]), 2) else [Discount], Replacer.ReplaceValue, {"Discount"}),
        #"Восстановлено Quantity" = Table.ReplaceValue(#"Восстановлен Discount", each [Quantity], each if [Quantity] = null and [SalesAmount] <> null and [Discount] <> null then Number.Round([SalesAmount] / ([UnitPrice] * (1 - [Discount]))) else [Quantity], Replacer.ReplaceValue, {"Quantity"}),
        #"Восстановлена SalesAmount" = Table.ReplaceValue(#"Восстановлено Quantity", each [SalesAmount], each if [SalesAmount] = null and [Quantity] <> null and [Discount] <> null then [Quantity] * [UnitPrice] * (1 - [Discount]) else [SalesAmount], Replacer.ReplaceValue, {"SalesAmount"}),
        #"Объединенные запросы1" = Table.NestedJoin(#"Восстановлена SalesAmount", {"City"}, stg_Regions, {"City"}, "stg_Regions", JoinKind.LeftOuter),
        #"Развернутый элемент stg_Regions" = Table.ExpandTableColumn(#"Объединенные запросы1", "stg_Regions", {"Region"}, {"RegionFromDirectory"}),
        #"Region из справочника" = Table.ReplaceValue(#"Развернутый элемент stg_Regions", each [Region], each if [Region] = null then [RegionFromDirectory] else [Region], Replacer.ReplaceValue, {"Region"}),
        #"Замененное значение null в Manager" = Table.ReplaceValue(#"Region из справочника", null, "Не указан", Replacer.ReplaceValue, {"Manager"}),
        #"Замененное значение null в Status" = Table.ReplaceValue(#"Замененное значение null в Manager", null, "Unknown", Replacer.ReplaceValue, {"Status"}),
        #"Добавлен SalesLevel" = Table.AddColumn(#"Замененное значение null в Status", "SalesLevel", each if [SalesAmount] = null then null else if [SalesAmount] >= 500000 then "High" else if [SalesAmount] >= 100000 then "Medium" else "Low", type text),
        #"Добавлен OrderQuality" = Table.AddColumn(#"Добавлен SalesLevel", "OrderQuality", each if [Quantity] = null or [SalesAmount] = null then "Check" else if [Quantity] <= 0 then "Check" else if [SalesAmount] <= 0 then "Check" else "OK", type text),
        #"Фамилия клиента (столбец из примеров)" = Table.AddColumn(#"Добавлен OrderQuality", "CustomerLastName", each Text.AfterDelimiter([CustomerName], " "), type text),
        #"Удаленные столбцы" = Table.RemoveColumns(#"Фамилия клиента (столбец из примеров)",{"StandardPrice", "RegionFromDirectory"}),
        #"Измененный тип" = Table.TransformColumnTypes(#"Удаленные столбцы",{{"Quantity", Int64.Type}, {"UnitPrice", type number}, {"Discount", type number}, {"SalesAmount", type number}, {"Region", type text}, {"Manager", type text}, {"Status", type text}})
    in
        #"Измененный тип"''', "10_MODEL",
    SALES_COLS + [("SalesLevel", "text"), ("OrderQuality", "text"), ("CustomerLastName", "text")],
    measures={"Total Sales": ("SUM(FactSales[SalesAmount])", "#,0"), "Rows": "COUNTROWS(FactSales)",
              "Check Rows": 'CALCULATE(COUNTROWS(FactSales), FactSales[OrderQuality] = "Check")'})

dim_product = Query("DimProduct", '''
    let
        Источник = stg_Products
    in
        Источник''', "10_MODEL",
    [("ProductCode", "text"), ("ProductName", "text"), ("Category", "text"), ("StandardPrice", "int"),
     ("SupplierCode", "text"), ("ProductGroupCode", "text"), ("ProductNumber", "text"), ("ProductTypeFromExample", "text")])

dim_region = Query("DimRegion", '''
    let
        Источник = stg_Regions
    in
        Источник''', "10_MODEL", [("Region", "text"), ("City", "text"), ("RegionalManager", "text")])

monthly = Query("SalesMonthly", '''
    let
        Источник = stg_MonthlySales,
        #"Отмененное свертывание других столбцов" = Table.UnpivotOtherColumns(Источник, {"Product"}, "Attribute", "Value"),
        #"Переименованные столбцы" = Table.RenameColumns(#"Отмененное свертывание других столбцов",{{"Attribute", "Month"}, {"Value", "Sales"}}),
        #"Добавлен MonthNumber" = Table.AddColumn(#"Переименованные столбцы", "MonthNumber", each List.PositionOf({"Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"}, [Month]) + 1, Int64.Type),
        #"Измененный тип" = Table.TransformColumnTypes(#"Добавлен MonthNumber",{{"Sales", Int64.Type}})
    in
        #"Измененный тип"''', "10_MODEL", [("Product", "text"), ("Month", "text"), ("Sales", "int"), ("MonthNumber", "int")])

by_city = Query("Sales_By_City", '''
    let
        Источник = FactSales,
        #"Сгруппированные строки" = Table.Group(Источник, {"City"}, {{"TotalSales", each List.Sum([SalesAmount]), type nullable number}, {"AvgSales", each List.Average([SalesAmount]), type nullable number}, {"RowsCount", each Table.RowCount(_), Int64.Type}, {"MaxSales", each List.Max([SalesAmount]), type nullable number}}),
        #"Сортированные строки" = Table.Sort(#"Сгруппированные строки",{{"TotalSales", Order.Descending}})
    in
        #"Сортированные строки"''', "20_ANALYSIS",
    [("City", "text"), ("TotalSales", "num"), ("AvgSales", "num"), ("RowsCount", "int"), ("MaxSales", "num")])

pivot = Query("Sales_City_Payment", '''
    let
        Источник = FactSales,
        #"Другие удаленные столбцы" = Table.SelectColumns(Источник,{"City", "PaymentMethod", "SalesAmount"}),
        #"Сводный столбец" = Table.Pivot(#"Другие удаленные столбцы", List.Distinct(#"Другие удаленные столбцы"[PaymentMethod]), "PaymentMethod", "SalesAmount", List.Sum)
    in
        #"Сводный столбец"''', "20_ANALYSIS",
    [("City", "text"), ("Card", "num"), ("Cash", "num"), ("Kaspi QR", "num"), ("QR", "num")])

queries = [src, stg_sales, stg_products, stg_regions, stg_monthly, fact, dim_product, dim_region, monthly, by_city, pivot]

if DEMO:  # задания 12 и 23: временные запросы, после скриншотов удаляются
    queries += [
        Query("TEMP_FillUp", '\n            let' + sheet("Regions") + '''
                #"Измененный тип" = Table.TransformColumnTypes(#"Удаленные пустые строки",{{"Region", type text}, {"City", type text}, {"RegionalManager", type text}}),
                #"Заполнено вверх" = Table.FillUp(#"Измененный тип",{"Region", "RegionalManager"})
            in
                #"Заполнено вверх"''', "99_TEST", load=False),
        Query("TEST_Sales_Duplicate", "\n            let" + STG_SALES_STEPS + ''',
                #"Удаленные дубликаты" = Table.Distinct(#"Замененные ошибки"),
                #"Удаленные столбцы" = Table.RemoveColumns(#"Удаленные дубликаты",{"Manager"})
            in
                #"Удаленные столбцы"''', "99_TEST", load=False),
    ]

build(BASE, queries,
      relationships=["FactSales[ProductCode] DimProduct[ProductCode]".split(), "FactSales[City] DimRegion[City]".split()],
      params={"pSourceFile": SRC}, sort_by={("SalesMonthly", "Month"): "MonthNumber"})

sales = proj(measure("FactSales", "Total Sales"))
write_pages(BASE + ".Report", [("ETL Results", [
    textbox((20, 10, 1880, 80), "Week 2 — Power Query ETL: Sales after cleaning", 26),
    visual("card", (20, 100, 440, 170), {"Values": [sales]}, "Total Sales"),
    visual("card", (480, 100, 440, 170), {"Values": [proj(measure("FactSales", "Rows"), "Rows after ETL")]}, "Rows after ETL"),
    visual("card", (940, 100, 440, 170), {"Values": [proj(measure("FactSales", "Check Rows"), "OrderQuality = Check")]}, "Rows to check"),
    visual("slicer", (1400, 100, 500, 170), {"Values": [proj(col("FactSales", "SalesLevel"))]}, "SalesLevel"),
    visual("clusteredBarChart", (20, 290, 620, 770), {"Category": [proj(col("Sales_By_City", "City"))],
                                                       "Y": [proj(agg("Sales_By_City", "TotalSales"), "TotalSales")]},
           "Group By: TotalSales by City", (agg("Sales_By_City", "TotalSales"), "Descending")),
    visual("tableEx", (660, 290, 1240, 360), {"Values": [proj(col("Sales_City_Payment", c)) for c in
                                                         ("City", "Card", "Cash", "Kaspi QR", "QR")]},
           "Pivot: City × PaymentMethod (Sum SalesAmount)"),
    visual("clusteredColumnChart", (660, 670, 1240, 390), {"Category": [proj(col("SalesMonthly", "Month"))],
                                                            "Y": [proj(agg("SalesMonthly", "Sales"), "Sales")]},
           "Unpivot: Monthly Sales (Product | Month | Sales)", (col("SalesMonthly", "Month"), "Ascending")),
])])
print("ok", "demo" if DEMO else "")
