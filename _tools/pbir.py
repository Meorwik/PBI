"""Мини-генератор визуалов PBIR (формат проекта Power BI .pbip).

Power BI должен быть закрыт, пока пишем файлы: при сохранении он их перезапишет.
"""
import json, os, shutil, uuid


def col(t, c):
    return {"Column": {"Expression": {"SourceRef": {"Entity": t}}, "Property": c}}


def measure(t, m):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": t}}, "Property": m}}


def agg(t, c, fn=0):  # 0 = Sum
    return {"Aggregation": {"Expression": col(t, c), "Function": fn}}


def lit(v):
    if isinstance(v, bool):
        v = "true" if v else "false"
    elif isinstance(v, (int, float)):
        v = f"{v}D"
    else:
        v = "'" + v.replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": v}}}


def proj(field, name=None):
    """Поле для роли визуала. name — подпись поля в этом визуале."""
    if "Aggregation" in field:
        a = field["Aggregation"]["Expression"]["Column"]
        ref = f"Sum({a['Expression']['SourceRef']['Entity']}.{a['Property']})"
    else:
        k = next(iter(field))
        ref = f"{field[k]['Expression']['SourceRef']['Entity']}.{field[k]['Property']}"
    p = {"field": field, "queryRef": ref, "nativeQueryRef": ref.split(".", 1)[1]}
    if name:
        p["displayName"] = name
    return p


def visual(vtype, pos, roles=None, title=None, sort=None, objects=None):
    """pos = (x, y, w, h) в пикселях страницы; roles = {"Category": [proj...], ...}."""
    x, y, w, h = pos
    v = {"visualType": vtype}
    if roles:
        v["query"] = {"queryState": {r: {"projections": ps} for r, ps in roles.items()}}
        if sort:
            field, direction = sort
            v["query"]["sortDefinition"] = {"sort": [{"field": field, "direction": direction}], "isDefaultSort": False}
    if objects:
        v["objects"] = objects
    if title:
        v["visualContainerObjects"] = {"title": [{"properties": {"show": lit(True), "text": lit(title)}}]}
    v["drillFilterOtherVisuals"] = True
    return {"position": {"x": x, "y": y, "z": 0, "height": h, "width": w}, "visual": v}


def textbox(pos, text, size=24):
    x, y, w, h = pos
    return {"position": {"x": x, "y": y, "z": 0, "height": h, "width": w},
            "visual": {"visualType": "textbox", "objects": {"general": [{"properties": {"paragraphs": [
                {"textRuns": [{"value": text, "textStyle": {"fontWeight": "bold", "fontSize": f"{size}pt"}}]}]}}]},
                "drillFilterOtherVisuals": True}}


def write_pages(report_dir, pages):
    """pages = [(displayName, [visual...]), ...]. Все старые страницы удаляются."""
    root = os.path.join(report_dir, "definition", "pages")
    shutil.rmtree(root, ignore_errors=True)
    order = []
    for i, (title, visuals) in enumerate(pages):
        pid = f"page{i + 1:02d}"
        order.append(pid)
        pdir = os.path.join(root, pid)
        os.makedirs(pdir)
        _dump(os.path.join(pdir, "page.json"), {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
            "name": pid, "displayName": title, "displayOption": "FitToPage", "height": 1080, "width": 1920})
        for z, vis in enumerate(visuals):
            vid = uuid.uuid4().hex[:20]
            vis = {"$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.13.0/schema.json",
                   "name": vid, **vis}
            vis["position"] = {**vis["position"], "z": z * 1000, "tabOrder": z * 1000}
            os.makedirs(os.path.join(pdir, "visuals", vid))
            _dump(os.path.join(pdir, "visuals", vid, "visual.json"), vis)
    _dump(os.path.join(root, "pages.json"), {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.1.0/schema.json",
        "pageOrder": order, "activePageName": order[0]})


def layout(model_dir, places):
    """Расставить таблицы в Model View: places = {"Sales": (x, y), ...}."""
    path = os.path.join(model_dir, "diagramLayout.json")
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    for node in d["diagrams"][0]["nodes"]:
        if node["nodeIndex"] in places:
            x, y = places[node["nodeIndex"]]
            node["location"] = {"x": x, "y": y}
    _dump(path, d)


def _dump(path, obj):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    assert proj(agg("Sales", "Sales Amount"))["queryRef"] == "Sum(Sales.Sales Amount)"
    assert proj(col("Product", "Category"))["queryRef"] == "Product.Category"
    assert lit("it's")["expr"]["Literal"]["Value"] == "'it''s'"
    print("ok")
