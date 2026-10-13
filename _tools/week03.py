"""Неделя 3: Append/Merge, папка, pFolderPath, функции, Left Anti.

python3 week03.py        — итоговый проект
python3 week03.py demo   — плюс тест неуникального ключа (СРСП A2), удаляется после скриншота
"""
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from pbip import Query, build
from pbir import agg, col, measure, proj, textbox, visual, write_pages

BASE = os.path.expanduser("~/PowerBI_2026/Week03/_pbip/Week3_Polienko")
DEMO = "demo" in sys.argv[1:]
P, S, ST, K, M = "00_Параметры", "01_Служебные", "02_Staging", "03_Контроль", "04_Модель"

SALES_TYPES = ('{{"Код_заказа", type text}, {"Дата", type date}, {"Код_товара", type text}, {"Код_магазина", type text}, '
               '{"Количество", Int64.Type}, {"Цена", type number}, {"Скидка", type number}, {"Канал", type text}, '
               '{"Менеджер", type text}, {"Статус", type text}}')
SALES_COLS = [("Код_заказа", "text"), ("Дата", "date"), ("Код_товара", "text"), ("Код_магазина", "text"),
              ("Количество", "int"), ("Цена", "num"), ("Скидка", "num"), ("Канал", "text"), ("Менеджер", "text"),
              ("Статус", "text")]


def month_file(name, file):
    return Query(name, f'''
        let
            Источник = Excel.Workbook(File.Contents(pFolderPath & "{file}"), null, true),
            SalesTable_Table = Источник{{[Item="SalesTable",Kind="Table"]}}[Data],
            #"Измененный тип" = Table.TransformColumnTypes(SalesTable_Table,{SALES_TYPES})
        in
            #"Измененный тип"''', ST, load=False)


def dim(name, item, types, cols):
    return Query(name, f'''
        let
            Источник = Excel.Workbook(File.Contents(pReferenceFile), null, true),
            {item}_Table = Источник{{[Item="{item}",Kind="Table"]}}[Data],
            #"Измененный тип" = Table.TransformColumnTypes({item}_Table,{types})
        in
            #"Измененный тип"''', ST, cols)


queries = [
    # --- то, что создаёт «Объединить файлы» ---
    Query("Пример файла", '''
        let
            Источник = Folder.Files(pFolderPath),
            Навигация1 = Источник{0}[Content]
        in
            Навигация1''', S, load=False, result="Binary"),
    Query("Параметр1", '#"Пример файла" meta [IsParameterQuery=true, BinaryIdentifier=#"Пример файла", Type="Binary", IsParameterQueryRequired=true]',
          S, load=False, result="Binary"),
    Query("Преобразовать пример файла", '''
        let
            Источник = Excel.Workbook(Параметр1, null, true),
            SalesTable_Table = Источник{[Item="SalesTable",Kind="Table"]}[Data]
        in
            SalesTable_Table''', S, load=False),
    Query("Преобразовать файл", '''
        let
            Источник = (Параметр1 as binary) => let
                    Источник = Excel.Workbook(Параметр1, null, true),
                    SalesTable_Table = Источник{[Item="SalesTable",Kind="Table"]}[Data]
                in
                    SalesTable_Table
        in
            Источник''', S, load=False, result="Function",
          attrs='[ FunctionQueryBinding = "{""exemplarFormulaName"":""Преобразовать пример файла""}" ]'),
    # --- свои функции ---
    Query("fnСуммаПослеСкидки", '''
        (qty as number, price as number, discount as number) as number =>
            Number.Round(qty * price * (1 - discount), 2)''', S, load=False, result="Function"),
    Query("fnТипОперации", '''
        (qty as number) as text =>
            if qty < 0 then "Возврат" else "Продажа"''', S, load=False, result="Function"),
    # --- staging ---
    Query("stg_ФайлыПродаж", f'''
        let
            Источник = Folder.Files(pFolderPath),
            #"Строки с примененным фильтром" = Table.SelectRows(Источник, each [Extension] = ".xlsx" and not Text.StartsWith([Name], "~$")),
            #"Вызвать настраиваемую функцию1" = Table.AddColumn(#"Строки с примененным фильтром", "Преобразовать файл", each #"Преобразовать файл"([Content])),
            #"Переименованные столбцы1" = Table.RenameColumns(#"Вызвать настраиваемую функцию1", {{"Name", "Source.Name"}}),
            #"Другие удаленные столбцы1" = Table.SelectColumns(#"Переименованные столбцы1", {{"Source.Name", "Преобразовать файл"}}),
            #"Развернутый столбец таблицы1" = Table.ExpandTableColumn(#"Другие удаленные столбцы1", "Преобразовать файл", Table.ColumnNames(#"Преобразовать файл"(#"Пример файла"))),
            #"Измененный тип" = Table.TransformColumnTypes(#"Развернутый столбец таблицы1",{{{{"Source.Name", type text}}, {SALES_TYPES[1:]})
        in
            #"Измененный тип"''', ST, load=False),
    dim("dim_Товары", "Products", '{{"Код_товара", type text}, {"Товар", type text}, {"Категория", type text}, {"Бренд", type text}, {"Базовая_цена", type number}}',
        [("Код_товара", "text"), ("Товар", "text"), ("Категория", "text"), ("Бренд", "text"), ("Базовая_цена", "num")]),
    dim("dim_Магазины", "Stores", '{{"Код_магазина", type text}, {"Город", type text}, {"Регион", type text}, {"Торговая_точка", type text}}',
        [("Код_магазина", "text"), ("Город", "text"), ("Регион", "text"), ("Торговая_точка", "text")]),
    month_file("stg_Январь", "Продажи_01_2026.xlsx"),
    month_file("stg_Февраль", "Продажи_02_2026.xlsx"),
    # --- уровень 1: ручные Append / Merge ---
    Query("test_Append_ЯнФев", '''
        let
            Источник = Table.Combine({stg_Январь, stg_Февраль})
        in
            Источник''', K, load=False),
    Query("test_Merge_Товары", '''
        let
            Источник = Table.NestedJoin(test_Append_ЯнФев, {"Код_товара"}, dim_Товары, {"Код_товара"}, "dim_Товары", JoinKind.LeftOuter),
            #"Развернутый элемент dim_Товары" = Table.ExpandTableColumn(Источник, "dim_Товары", {"Товар", "Категория", "Бренд"}, {"Товар", "Категория", "Бренд"})
        in
            #"Развернутый элемент dim_Товары"''', K, load=False),
    # --- рабочий запрос ---
    Query("Продажи", '''
        let
            Источник = stg_ФайлыПродаж,
            #"Вызвана пользовательская функция" = Table.AddColumn(Источник, "Сумма_после_скидки", each fnСуммаПослеСкидки([Количество], [Цена], [Скидка])),
            #"Вызвана пользовательская функция1" = Table.AddColumn(#"Вызвана пользовательская функция", "Контроль_Статуса", each fnТипОперации([Количество])),
            #"Объединенные запросы" = Table.NestedJoin(#"Вызвана пользовательская функция1", {"Код_товара"}, dim_Товары, {"Код_товара"}, "dim_Товары", JoinKind.LeftOuter),
            #"Развернутый элемент dim_Товары" = Table.ExpandTableColumn(#"Объединенные запросы", "dim_Товары", {"Товар", "Категория", "Бренд"}, {"Товар", "Категория", "Бренд"}),
            #"Объединенные запросы1" = Table.NestedJoin(#"Развернутый элемент dim_Товары", {"Код_магазина"}, dim_Магазины, {"Код_магазина"}, "dim_Магазины", JoinKind.LeftOuter),
            #"Развернутый элемент dim_Магазины" = Table.ExpandTableColumn(#"Объединенные запросы1", "dim_Магазины", {"Город", "Регион", "Торговая_точка"}, {"Город", "Регион", "Торговая_точка"}),
            #"Измененный тип" = Table.TransformColumnTypes(#"Развернутый элемент dim_Магазины",{{"Сумма_после_скидки", type number}, {"Контроль_Статуса", type text}})
        in
            #"Измененный тип"''', M,
          [("Source.Name", "text")] + SALES_COLS + [("Сумма_после_скидки", "num"), ("Контроль_Статуса", "text"),
           ("Товар", "text"), ("Категория", "text"), ("Бренд", "text"), ("Город", "text"), ("Регион", "text"),
           ("Торговая_точка", "text")],
          measures={"Строк": "COUNTROWS('Продажи')", "Файлов": "DISTINCTCOUNT('Продажи'[Source.Name])",
                    "Выручка после скидки": ("SUM('Продажи'[Сумма_после_скидки])", "#,0"),
                    "Неизвестных товаров": "COUNTROWS('dq_НеизвестныеТовары')",
                    "Неизвестных магазинов": "COUNTROWS('dq_НеизвестныеМагазины')",
                    "Расхождений статуса": "COUNTROWS('dq_РасхожденияСтатуса') + 0"}),
    # --- контроль качества ---
    Query("dq_НеизвестныеТовары", '''
        let
            Источник = Table.NestedJoin(stg_ФайлыПродаж, {"Код_товара"}, dim_Товары, {"Код_товара"}, "dim_Товары", JoinKind.LeftAnti),
            #"Удаленные столбцы" = Table.RemoveColumns(Источник,{"dim_Товары"})
        in
            #"Удаленные столбцы"''', K, [("Source.Name", "text")] + SALES_COLS),
    Query("dq_НеизвестныеМагазины", '''
        let
            Источник = Table.NestedJoin(stg_ФайлыПродаж, {"Код_магазина"}, dim_Магазины, {"Код_магазина"}, "dim_Магазины", JoinKind.LeftAnti),
            #"Удаленные столбцы" = Table.RemoveColumns(Источник,{"dim_Магазины"})
        in
            #"Удаленные столбцы"''', K, [("Source.Name", "text")] + SALES_COLS),
    Query("dq_РасхожденияСтатуса", '''
        let
            Источник = Продажи,
            #"Строки с примененным фильтром" = Table.SelectRows(Источник, each [Контроль_Статуса] <> [Статус]),
            #"Другие удаленные столбцы" = Table.SelectColumns(#"Строки с примененным фильтром",{"Source.Name", "Код_заказа", "Количество", "Статус", "Контроль_Статуса"})
        in
            #"Другие удаленные столбцы"''', K,
          [("Source.Name", "text"), ("Код_заказа", "text"), ("Количество", "int"), ("Статус", "text"), ("Контроль_Статуса", "text")]),
]

if DEMO:  # СРСП A2: неуникальный ключ справочника размножает строки
    queries += [
        Query("TEST_dim_Товары_дубль", '''
            let
                Источник = dim_Товары,
                #"Добавленный запрос" = Table.Combine({Источник, Table.SelectRows(dim_Товары, each [Код_товара] = "P001")})
            in
                #"Добавленный запрос"''', K, load=False),
        Query("TEST_Merge_дубль", '''
            let
                Источник = Table.NestedJoin(test_Append_ЯнФев, {"Код_товара"}, TEST_dim_Товары_дубль, {"Код_товара"}, "dim", JoinKind.LeftOuter),
                #"Развернутый элемент dim" = Table.ExpandTableColumn(Источник, "dim", {"Товар"}, {"Товар"})
            in
                #"Развернутый элемент dim"''', K, load=False),
    ]

build(BASE, queries, params={"pFolderPath": r"Z:\Week03\01_Продажи_по_месяцам" + "\\",
                             "pReferenceFile": r"Z:\Week03\Справочники_и_план.xlsx"}, param_group=P)

m = lambda name: proj(measure("Продажи", name))
write_pages(BASE + ".Report", [("Week 3 — Folder ETL", [
    textbox((20, 10, 1880, 80), "Неделя 3 — Папка, параметр, функции, контроль ключей", 26),
    visual("card", (20, 100, 360, 160), {"Values": [m("Строк")]}, "Строк продаж (Folder)"),
    visual("card", (395, 100, 360, 160), {"Values": [m("Файлов")]}, "Файлов в папке"),
    visual("card", (770, 100, 380, 160), {"Values": [m("Выручка после скидки")]}, "fnСуммаПослеСкидки, итог"),
    visual("card", (1165, 100, 360, 160), {"Values": [m("Неизвестных товаров")]}, "Left Anti: товары"),
    visual("card", (1540, 100, 360, 160), {"Values": [m("Неизвестных магазинов")]}, "Left Anti: магазины"),
    visual("clusteredColumnChart", (20, 280, 900, 380), {"Category": [proj(col("Продажи", "Source.Name"))], "Y": [m("Строк")]},
           "Строк по исходным файлам (Source.Name)", (col("Продажи", "Source.Name"), "Ascending")),
    visual("clusteredBarChart", (940, 280, 960, 380), {"Category": [proj(col("Продажи", "Категория"))], "Y": [m("Выручка после скидки")]},
           "Выручка по категориям (Merge с dim_Товары)", (measure("Продажи", "Выручка после скидки"), "Descending")),
    visual("tableEx", (20, 680, 900, 380), {"Values": [proj(col("dq_НеизвестныеТовары", c)) for c in
                                                       ("Source.Name", "Код_заказа", "Код_товара", "Код_магазина", "Количество")]},
           "dq_НеизвестныеТовары (Left Anti)"),
    visual("tableEx", (940, 680, 960, 380), {"Values": [proj(col("dq_НеизвестныеМагазины", c)) for c in
                                                        ("Source.Name", "Код_заказа", "Код_товара", "Код_магазина", "Количество")]},
           "dq_НеизвестныеМагазины (Left Anti)"),
])])
print("ok", "demo" if DEMO else "")
