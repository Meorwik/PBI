"""Сборка проекта Power BI (.pbip) из кода: запросы M, таблицы, связи, меры.

Каркас отчёта (тема, report.json) берём из Week01 — его сохранил сам Power BI.
Открыть .pbip в Power BI → «Обновить» → «Сохранить как» .pbix.
"""
import json, os, shutil, textwrap, uuid

TEMPLATE = os.path.expanduser("~/PowerBI_2026/Week01/_pbip/Week01_Polienko_Ilya.Report")
TYPES = {"text": "string", "int": "int64", "num": "double", "date": "dateTime", "bool": "boolean"}


class Query:
    """M-запрос. cols = [(имя, тип)] только для загружаемых в модель (load=True)."""

    def __init__(self, name, m, group=None, cols=None, load=True, measures=None, hidden=(), result="Table", attrs=None, fmt=None, calc=None):
        self.name, self.m, self.group, self.cols = name, textwrap.dedent(m).strip(), group, cols or []
        self.load, self.measures, self.hidden = load, measures or {}, set(hidden)
        self.result, self.attrs, self.fmt = result, attrs, fmt or {}
        self.calc = calc or {}  # вычисляемые столбцы: {имя: (DAX, тип, формат)}  # result: Table/Function/Binary; attrs — mAttributes (Combine Files)


def q(name):
    """Имя в TMDL: в кавычках, если есть пробелы/спецсимволы."""
    return name if name.isascii() and name.replace("_", "").isalnum() and not name[0].isdigit() else "'" + name.replace("'", "''") + "'"


def _indent(text, n):
    return "\n".join("\t" * n + line if line else "" for line in text.splitlines())


def build(base, queries, relationships=(), params=None, sort_by=None, param_group="Параметры", layout=None):
    """layout = {таблица: (x, y)} — расположение в Model View."""
    """base = .../Name (без расширения); params = {имя: значение-текст}; sort_by = {(таблица, столбец): столбец}."""
    name = os.path.basename(base)
    sm, rep = base + ".SemanticModel", base + ".Report"
    shutil.rmtree(sm, ignore_errors=True)
    d = os.path.join(sm, "definition")
    os.makedirs(os.path.join(d, "tables"))
    os.makedirs(os.path.join(d, "cultures"))
    _w(os.path.join(sm, "definition.pbism"), json.dumps({"version": "4.2", "settings": {}}, indent=2))
    _w(os.path.join(sm, ".platform"), _platform("SemanticModel", name))
    _w(os.path.join(d, "database.tmdl"), "database\n\tcompatibilityLevel: 1606\n\n")
    _w(os.path.join(d, "cultures", "ru-RU.tmdl"), "cultureInfo ru-RU\n\n")
    groups = sorted({x.group for x in queries if x.group} | ({param_group} if params else set()))
    loaded = [x for x in queries if x.load]
    model = ["model Model", "\tculture: ru-RU", "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
             "\tsourceQueryCulture: ru-RU", "\tdataAccessOptions", "\t\tlegacyRedirects", "\t\treturnErrorValuesAsNull", ""]
    for i, g in enumerate(groups):
        model += [f"queryGroup {q(g)}", "", f"\tannotation PBI_QueryGroupOrder = {i}", ""]
    model += ["annotation __PBI_TimeIntelligenceEnabled = 0", "",
              "annotation PBI_QueryOrder = " + json.dumps([*(params or {}), *(x.name for x in queries)], ensure_ascii=False), ""]
    model += [f"ref table {q(x.name)}" for x in loaded] + ["", "ref cultureInfo ru-RU", ""]
    _w(os.path.join(d, "model.tmdl"), "\n".join(model))

    expr = []
    for p, v in (params or {}).items():
        expr += [f'expression {q(p)} = "{v}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]',
                 f"\tlineageTag: {uuid.uuid4()}", f"\tqueryGroup: {q(param_group)}", "", "\tannotation PBI_ResultType = Text", ""]
    for x in queries:
        if x.load:
            continue
        expr += [f"expression {q(x.name)} =", _indent(x.m, 2)]
        if x.attrs:
            expr.append(f"\tmAttributes: {x.attrs}")
        expr.append(f"\tlineageTag: {uuid.uuid4()}")
        if x.group:
            expr.append(f"\tqueryGroup: {q(x.group)}")
        expr += ["", f"\tannotation PBI_ResultType = {x.result}", ""]
    if expr:
        _w(os.path.join(d, "expressions.tmdl"), "\n".join(expr))

    tags = {}
    for x in loaded:
        tags[x.name] = str(uuid.uuid4())
        t = [f"table {q(x.name)}", f"\tlineageTag: {tags[x.name]}", ""]
        for mname, dax in x.measures.items():
            fmt = folder = None
            if isinstance(dax, tuple):
                dax, fmt, folder = (*dax, None)[:3]
            if "\n" in dax:  # многострочная формула: тело с отступом, как сохраняет сам Power BI
                t += [f"\tmeasure {q(mname)} =", _indent(textwrap.dedent(dax).strip(), 3)]
            else:
                t += [f"\tmeasure {q(mname)} = {dax}"]
            if fmt:
                t.append(f"\t\tformatString: {fmt}")
            if folder:
                t.append(f"\t\tdisplayFolder: {folder}")
            t += [f"\t\tlineageTag: {uuid.uuid4()}", ""]
        for cname, (dax, typ, fmt) in x.calc.items():
            t += [f"\tcolumn {q(cname)} = {dax}", f"\t\tdataType: {TYPES[typ]}"]
            if fmt:
                t.append(f"\t\tformatString: {fmt}")
            if cname in x.hidden:
                t.append("\t\tisHidden")
            t += [f"\t\tlineageTag: {uuid.uuid4()}", "\t\tsummarizeBy: sum", "", "\t\tannotation SummarizationSetBy = Automatic", ""]
        for c, typ in x.cols:
            t += [f"\tcolumn {q(c)}", f"\t\tdataType: {TYPES[typ]}"]
            if c in x.fmt:
                t.append(f"\t\tformatString: {x.fmt[c]}")
            elif typ == "date":
                t.append("\t\tformatString: Short Date")
            elif typ == "int":
                t.append("\t\tformatString: 0")
            if c in x.hidden:
                t.append("\t\tisHidden")
            additive = typ in ("int", "num") and not c.endswith(("ID", "Key", "Number"))
            t += [f"\t\tlineageTag: {uuid.uuid4()}", f"\t\tsummarizeBy: {'sum' if additive else 'none'}",
                  f"\t\tsourceColumn: {c}"]
            if sort_by and (x.name, c) in sort_by:
                t.append(f"\t\tsortByColumn: {q(sort_by[(x.name, c)])}")
            t += ["", "\t\tannotation SummarizationSetBy = Automatic", ""]
        t += [f"\tpartition {q(x.name)} = m", "\t\tmode: import"]
        if x.group:
            t.append(f"\t\tqueryGroup: {q(x.group)}")
        t += ["\t\tsource =", _indent(x.m, 4), "", "\tannotation PBI_ResultType = Table", ""]
        _w(os.path.join(d, "tables", f"{x.name}.tmdl"), "\n".join(t))

    rels = []
    for r in relationships:
        frm, to, *opts = r
        rels += [f"relationship {uuid.uuid4()}", f"\tfromColumn: {_col(frm)}", f"\ttoColumn: {_col(to)}"]
        for o in opts:  # например "isActive: false", "crossFilteringBehavior: bothDirections"
            rels.append("\t" + o)
        rels.append("")
    if rels:
        _w(os.path.join(d, "relationships.tmdl"), "\n".join(rels))

    if layout:
        nodes = [{"location": {"x": xy[0], "y": xy[1]}, "nodeIndex": n, "nodeLineageTag": tags[n],
                  "size": {"height": 300, "width": 234}, "zIndex": 0} for n, xy in layout.items()]
        _w(os.path.join(sm, "diagramLayout.json"), json.dumps({"version": "1.1.0", "diagrams": [{
            "ordinal": 0, "scrollPosition": {"x": 0, "y": 0}, "nodes": nodes, "name": "Все таблицы", "zoomValue": 100,
            "pinKeyFieldsToTop": False, "showExtraHeaderInfo": False, "hideKeyFieldsWhenCollapsed": False,
            "tablesLocked": False}], "selectedDiagram": "Все таблицы", "defaultDiagram": "Все таблицы"},
            ensure_ascii=False, indent=2))

    if not os.path.exists(rep):
        shutil.copytree(TEMPLATE, rep, ignore=shutil.ignore_patterns(".pbi", "pages"))
        _w(os.path.join(rep, "definition.pbir"), json.dumps(
            {"version": "4.0", "datasetReference": {"byPath": {"path": f"../{name}.SemanticModel"}}}, indent=2))
        _w(os.path.join(rep, ".platform"), _platform("Report", name))
    _w(base + ".pbip", json.dumps({"version": "1.0", "artifacts": [{"report": {"path": f"{name}.Report"}}],
                                   "settings": {"enableAutoRecovery": True}}, indent=2))


def _col(ref):
    t, c = ref.split("[")
    return f"{q(t)}.{q(c.rstrip(']'))}"


def _platform(kind, name):
    return json.dumps({"$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
                       "metadata": {"type": kind, "displayName": name},
                       "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}}, indent=2)


def _w(path, text):
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)


if __name__ == "__main__":
    assert q("Продажи") == "'Продажи'" and q("FactSales") == "FactSales" and q("Kaspi QR") == "'Kaspi QR'" and q("00_STAGING") == "'00_STAGING'"
    assert _col("FactSales[City]") == "FactSales.City"
    print("ok")
