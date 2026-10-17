"""Keep only parts reachable from the selected, static print workbook."""

from io import BytesIO
import zipfile
import xml.etree.ElementTree as ET


def print_package(parts: dict[str, bytes], selected_id: str) -> bytes:
    from app.excel import xml, related, relationship_file
    workbook_rels = "xl/_rels/workbook.xml.rels"
    root = xml(parts[workbook_rels])
    for relation in list(root):
        kind = relation.get("Type", "").rsplit("/", 1)[-1]
        if kind not in {"worksheet", "styles", "sharedStrings", "theme"} or (
                kind == "worksheet" and relation.get("Id") != selected_id):
            root.remove(relation)
    parts[workbook_rels] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    retained, pending = set(), ["xl/workbook.xml"]
    if "_rels/.rels" in parts:
        retained.add("_rels/.rels")
        for relation in xml(parts["_rels/.rels"]):
            if relation.get("TargetMode") != "External":
                pending.append(related("root.xml", relation.get("Target", "")))
    while pending:
        part = pending.pop()
        if part in retained or part not in parts:
            continue
        retained.add(part)
        path = relationship_file(part)
        if path not in parts:
            continue
        retained.add(path)
        relations = xml(parts[path])
        for relation in list(relations):
            kind = relation.get("Type", "").rsplit("/", 1)[-1]
            if relation.get("TargetMode") == "External" or kind in {"printerSettings", "calcChain", "oleObject", "control", "vbaProject"}:
                relations.remove(relation)
            else:
                pending.append(related(part, relation.get("Target", "")))
        parts[path] = ET.tostring(relations, encoding="utf-8", xml_declaration=True)
    content = xml(parts["[Content_Types].xml"])
    for entry in list(content):
        name = entry.get("PartName", "").lstrip("/")
        if name and name not in retained:
            content.remove(entry)
        elif name == "xl/workbook.xml":
            entry.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    parts["[Content_Types].xml"] = ET.tostring(content, encoding="utf-8", xml_declaration=True)
    retained.add("[Content_Types].xml")
    output = BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as target:
        for name in sorted(retained):
            target.writestr(name, parts[name])
    return output.getvalue()
