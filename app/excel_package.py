"""Keep only parts reachable from the selected, static print workbook."""

from io import BytesIO
import zipfile
import xml.etree.ElementTree as ET


def print_package(parts: dict[str, bytes], selected_id: str) -> bytes:
    from app.excel import xml, related, relationship_file
    content = xml(parts["[Content_Types].xml"])
    active = {entry.get("PartName", "").lstrip("/") for entry in content
              if any(kind in entry.get("ContentType", "").lower()
                     for kind in ("vbaproject", "vbadata", "macrosheet", "activex"))}

    def blocked(part):
        return part in active or any(token in part.lower()
                                     for token in ("vbaproject", "/activex/", "/macrosheets/"))

    excluded = {"printerSettings", "calcChain", "oleObject", "control", "ctrlProp",
                "vbaProject", "vbaData", "xlMacrosheet", "xlIntlMacrosheet", "extensibility"}

    def destinations(part, root):
        for relation in list(root):
            target = related(part, relation.get("Target", "")) if relation.get("TargetMode") != "External" else ""
            kind = relation.get("Type", "").rsplit("/", 1)[-1]
            if not target or kind in excluded or blocked(target):
                root.remove(relation)
            else:
                yield target

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
        root = xml(parts["_rels/.rels"])
        pending.extend(destinations("root.xml", root))
        parts["_rels/.rels"] = ET.tostring(root, encoding="utf-8", xml_declaration=True)
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
        pending.extend(destinations(part, relations))
        parts[path] = ET.tostring(relations, encoding="utf-8", xml_declaration=True)
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
