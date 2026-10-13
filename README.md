# Power BI — практические работы (DS 24-Б, Полиенко Илья)

| Неделя | Тема | Файл Power BI | Отчёт |
|---|---|---|---|
| 1 | Установка, первый отчёт (AdventureWorks) | `Week01/Week01_Polienko_Ilya.pbix` | `Week01/Отчёт_Неделя1.docx` |
| 2 | Power Query ETL: очистка «грязных» продаж | `Week02/DS24-B_Polienko_Week2.pbix` | `Week02/Отчёт_Неделя2.docx` |
| 3 | Append/Merge, папка, параметры, функции, Left Anti | `Week03/Week3_Polienko.pbix` | `Week03/Отчёт_Неделя3.docx` ⚠️ |
| 4 | Flat Table → Star Schema | `Week04/DS24-B_Polienko_Week4.pbix` | `Week04/Отчёт_Неделя4.docx` |
| 5 | Связи: 2 факта, active/inactive, 1:1, M:M, Single/Both | `Week05/DS24-B_Polienko_Week5.pbix` | `Week05/Отчёт_Неделя5.docx` |
| 6 | DAX: _Measures, итераторы, контексты, VAR, IF/SWITCH | `Week06/DS24-B_Polienko_Week6.pbix` | `Week06/Отчёт_Неделя6.docx` |

⚠️ Неделя 3: в пакете не хватает файлов мая и июня. Без них не сделаны Refresh-тест 600 → 720 и schema drift. Всё остальное готово.

`СЦЕНАРИЙ_ВИДЕО.md` — что показывать и говорить на видео по каждой неделе.

В каждой папке:
- `*.pbix` — готовый файл;
- `Отчёт_*.md/.docx` — отчёт, СРСП и ответы на вопросы защиты;
- `Скриншоты/` — скриншоты по заданию;
- `_pbip/` — тот же проект в текстовом формате Power BI Project (TMDL + PBIR): удобно смотреть M-код и DAX.

`_tools/` — скрипты, которые собирают проекты `.pbip` из кода:
- `pbip.py` — запросы M, таблицы, связи, меры;
- `pbir.py` — визуалы;
- `weekNN.py` — содержимое конкретной недели.

Если `.pbix` открывается на другом компьютере, нужно поменять путь к Excel:
- недели 2, 4, 5, 6 — параметр `pSourceFile`;
- неделя 3 — параметры `pFolderPath` и `pReferenceFile`;
- неделя 1 — «Настройки источника данных».
