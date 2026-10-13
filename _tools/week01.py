"""Неделя 1: страница Sales Overview + связь Sales[OrderDateKey] → Date[DateKey]."""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from pbir import agg, col, layout, lit, proj, textbox, visual, write_pages

BASE = os.path.expanduser("~/PowerBI_2026/Week01/_pbip/Week01_Polienko_Ilya")
REL = """
relationship Sales_OrderDate
	fromColumn: Sales.OrderDateKey
	toColumn: Date.DateKey
"""

rel_path = BASE + ".SemanticModel/definition/relationships.tmdl"
with open(rel_path, encoding="utf-8") as f:
    rels = f.read()
if "Sales.OrderDateKey" not in rels:
    with open(rel_path, "a", encoding="utf-8") as f:
        f.write(REL)

# звезда: факт в центре, измерения по бокам
layout(BASE + ".SemanticModel", {"Sales": (480, 200), "Product": (120, 0), "Customer": (120, 380),
                                  "Date": (840, 0), "SalesTerritory": (840, 380)})

sales = proj(agg("Sales", "Sales Amount"), "Sales Amount")
qty = proj(agg("Sales", "Order Quantity"), "Order Quantity")
category = proj(col("Product", "Category"))
by_sales = (agg("Sales", "Sales Amount"), "Descending")

write_pages(BASE + ".Report", [("Sales Overview", [
    textbox((20, 10, 1880, 80), "AdventureWorks Sales Dashboard", 28),
    visual("card", (20, 100, 360, 180), {"Values": [proj(agg("Sales", "Sales Amount"), "Total Sales")]}, "Total Sales"),
    visual("card", (20, 300, 360, 180), {"Values": [proj(agg("Sales", "Order Quantity"), "Total Quantity")]}, "Total Quantity"),
    visual("slicer", (20, 500, 360, 560), {"Values": [category]}, "Product Category"),
    visual("clusteredColumnChart", (400, 100, 740, 460), {"Category": [category], "Y": [sales]},
           "Sales by Product Category", by_sales),
    # Самостоятельное задание, вариант A
    visual("clusteredBarChart", (1160, 100, 740, 460), {"Category": [proj(col("SalesTerritory", "Region"))], "Y": [sales]},
           "Sales by Territory", by_sales),
    visual("lineChart", (400, 580, 740, 480), {"Category": [proj(col("Date", "Month"))], "Y": [sales]}, "Sales Trend",
           objects={"lineStyles": [{"properties": {"lineChartType": lit("linear")}}]}),
    visual("tableEx", (1160, 580, 740, 480),
           {"Values": [category, proj(col("Product", "Product")), sales, qty]}, "Products", by_sales),
])])
print("ok")
