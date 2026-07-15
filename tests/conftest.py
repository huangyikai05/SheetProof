"""Shared factories for small, source-generated workbook fixtures."""

from __future__ import annotations

import posixpath
from collections.abc import Callable
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook

WorkbookConfigurator = Callable[[Workbook], None]
WorkbookFactory = Callable[[str, WorkbookConfigurator | None], Path]

_CONTENT_TYPES_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/content-types"
)
_RELATIONSHIPS_NAMESPACE = (
    "http://schemas.openxmlformats.org/package/2006/relationships"
)
_VBA_RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
)
_VBA_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
_MACRO_WORKBOOK_CONTENT_TYPE = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"


@pytest.fixture
def workbook_factory(tmp_path: Path) -> WorkbookFactory:
    """Create a workbook under pytest's temporary directory.

    Tests intentionally generate every OOXML input at runtime.  No opaque binary
    fixtures are checked into the repository.
    """

    def create(name: str, configure: WorkbookConfigurator | None = None) -> Path:
        path = tmp_path / name
        workbook = Workbook()
        if configure is not None:
            configure(workbook)
        workbook.save(path)
        workbook.close()
        return path

    return create


def add_vba_project(
    path: Path,
    payload: bytes = b"SheetProof test VBA marker",
    *,
    part_name: str = "xl/vbaProject.bin",
    relationship_target: str | None = None,
    content_type_part_name: str | None = None,
    content_type: str = _VBA_CONTENT_TYPE,
    include_relationship: bool = True,
    include_content_type: bool = True,
    content_type_as_default: bool = False,
    override_content_type: str | None = None,
    target_mode: str | None = None,
) -> None:
    """Add an inert VBA part and configurable OPC metadata for parser tests."""

    with ZipFile(path) as source:
        entries = [(info, source.read(info.filename)) for info in source.infolist()]
        archive_comment = source.comment

    content_types_tag = f"{{{_CONTENT_TYPES_NAMESPACE}}}Override"
    default_tag = f"{{{_CONTENT_TYPES_NAMESPACE}}}Default"
    relationship_tag = f"{{{_RELATIONSHIPS_NAMESPACE}}}Relationship"
    content_types_payload: bytes | None = None
    relationships_payload: bytes | None = None
    for info, data in entries:
        if info.filename == "[Content_Types].xml":
            root = ElementTree.fromstring(data)
            for element in root.findall(content_types_tag):
                if element.attrib.get("PartName", "").casefold() == "/xl/workbook.xml":
                    element.set("ContentType", _MACRO_WORKBOOK_CONTENT_TYPE)
                    break
            metadata_part_name = content_type_part_name or part_name
            package_part_name = f"/{metadata_part_name.lstrip('/')}"
            if include_content_type:
                if content_type_as_default:
                    extension = posixpath.basename(metadata_part_name).rsplit(".", 1)[-1]
                    ElementTree.SubElement(
                        root,
                        default_tag,
                        Extension=extension,
                        ContentType=content_type,
                    )
                else:
                    ElementTree.SubElement(
                        root,
                        content_types_tag,
                        PartName=package_part_name,
                        ContentType=content_type,
                    )
            if override_content_type is not None:
                ElementTree.SubElement(
                    root,
                    content_types_tag,
                    PartName=package_part_name,
                    ContentType=override_content_type,
                )
            content_types_payload = ElementTree.tostring(root, encoding="utf-8")
        elif info.filename == "xl/_rels/workbook.xml.rels" and include_relationship:
            root = ElementTree.fromstring(data)
            relationship_ids = {
                element.attrib.get("Id", "")
                for element in root.findall(relationship_tag)
            }
            relationship_index = 1
            while f"rId{relationship_index}" in relationship_ids:
                relationship_index += 1
            attributes = {
                "Id": f"rId{relationship_index}",
                "Type": _VBA_RELATIONSHIP_TYPE,
                "Target": relationship_target
                or posixpath.relpath(part_name.lstrip("/"), "xl"),
            }
            if target_mode is not None:
                attributes["TargetMode"] = target_mode
            ElementTree.SubElement(root, relationship_tag, attributes)
            relationships_payload = ElementTree.tostring(root, encoding="utf-8")

    rewritten = path.with_name(f"{path.stem}-with-vba{path.suffix}")
    with ZipFile(rewritten, "w", compression=ZIP_DEFLATED) as target:
        target.comment = archive_comment
        for info, data in entries:
            if info.filename == "[Content_Types].xml" and content_types_payload is not None:
                data = content_types_payload
            elif (
                info.filename == "xl/_rels/workbook.xml.rels"
                and relationships_payload is not None
            ):
                data = relationships_payload
            target.writestr(info, data)
        target.writestr(part_name, payload)
    rewritten.replace(path)
