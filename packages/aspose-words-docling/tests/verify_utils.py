# Copyright (c) 2001-2026 Aspose Pty Ltd.

import os
import pytest
import sys
import json

from pathlib import Path
from difflib import Differ
from typing import Any, List

from docling_core.types.doc import (
    DoclingDocument, 
    RefItem,
    DocItem,
    Formatting,
    Script
)
from docling.document_converter import DocumentConverter

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'src', 'aspose_words_docling')))
from aspose_words_converter import AsposeWordsConverter


TESTS_DIR = os.path.join(Path(__file__).parent, "test_files")
ARTIFACTS_DIR = os.path.join(Path(__file__).parent, "artifacts")
DIFF_TOOL = os.getenv("DIFF_TOOL_PATH", "X:\\Tools\\WinMerge\\WinMergeU.exe")
default_indent: str = "     "


def get_in_test_dir(filename: str) -> str:
    dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(TESTS_DIR, filename)


def get_artifacts_dir() -> str:
    return ARTIFACTS_DIR


def get_document_tree(doc: DoclingDocument) -> str:
    out_list: List[str] = []        
    _get_node_tree(doc, doc.body, 0, out_list)
    return "".join(out_list)


def _get_children_tree(doc: DoclingDocument, children: List[RefItem], level, out_list: List[str]):
    for item in children:
        _get_node_tree(doc, item.resolve(doc), level + 1, out_list)


def _get_node_tree(doc: DoclingDocument, node: Any, level: int, out_list: List[str]):
    max_text_len = 80
    ref_len = 9
    indent = default_indent * level
    self_ref = getattr(node, "self_ref", None)       
    label = getattr(node, "label", None)       
    text = _to_single_line_text(getattr(node, "text", ""))
    text = (text[:max_text_len] + '...') if len(text) > max_text_len else text
    text = text if len(text) == 0 else f" '{text}'"

    out_list.append(f"{indent}{(_strip_ref(self_ref) + ":").ljust(ref_len)} ({label}){text}\n")
    children = getattr(node, "children", [])
    _get_children_tree(doc, children, level, out_list)


def get_document_items(doc: DoclingDocument) -> str:
    out_list: List[str] = []
    _get_document_texts(doc, out_list)
    _get_document_groups(doc, out_list)
    _get_document_tables(doc, out_list)
    _get_document_pictures(doc, out_list)

    return "".join(out_list)

def _get_document_texts(doc: DoclingDocument, out_list: List[str]):
    for t in doc.texts:
        out_list.append(f"{_strip_ref(t.self_ref)}:\n")
        _add_parent_if_exists(t, out_list)
        out_list.append(f"{default_indent}label: {t.label}\n")
        _add_property_if_exists(t, "level", out_list)
        _add_property_if_exists(t, "name", out_list)
        out_list.append(f"{default_indent}text: {_to_single_line_text(t.text)}\n")
        out_list.append(f"{default_indent}orig: {t.orig}\n")
        out_list.append(f"{default_indent}formatting: {_get_formatting_string(t.formatting)}\n")
        _add_property_if_exists(t, "enumerated", out_list)
        _add_property_if_exists(t, "marker", out_list)
        out_list.append("\n")


def _get_document_groups(doc: DoclingDocument, out_list: List[str]):
    for g in doc.groups:
        out_list.append(f"{_strip_ref(g.self_ref)}:\n")
        _add_parent_if_exists(g, out_list)
        out_list.append(f"{default_indent}label: {g.label}\n")
        _add_property_if_exists(g, "level", out_list)
        _add_property_if_exists(g, "name", out_list)
        out_list.append("\n")


def _get_document_tables(doc: DoclingDocument, out_list: List[str]):
    for t in doc.tables:
        out_list.append(f"{_strip_ref(t.self_ref)}:\n")
        _add_parent_if_exists(t, out_list)
        cell_index = 0
        for c in t.data.table_cells:
            out_list.append(f"{default_indent}cell/{cell_index}:\n")
            _add_property_if_exists(c, "row_span", out_list, 2)
            _add_property_if_exists(c, "col_span", out_list, 2)
            _add_property_if_exists(c, "start_row_offset_idx", out_list, 2)
            _add_property_if_exists(c, "end_row_offset_idx", out_list, 2)
            _add_property_if_exists(c, "start_col_offset_idx", out_list, 2)
            _add_property_if_exists(c, "end_col_offset_idx", out_list, 2)
            _add_property_if_exists(c, "text", out_list, 2)
            _add_property_if_exists(c, "column_header", out_list, 2)
            _add_property_if_exists(c, "row_header", out_list, 2)
            _add_property_if_exists(c, "row_section", out_list, 2)
            cell_index += 1

        out_list.append("\n")


def _get_document_pictures(doc: DoclingDocument, out_list: List[str]):
    for p in doc.pictures:
        out_list.append(f"{_strip_ref(p.self_ref)}:\n")
        _add_parent_if_exists(p, out_list)
        out_list.append(f"{default_indent}label: {p.label}\n")
        _add_property_if_exists(p, "name", out_list)
        img = p.get_image(doc)
        base64_str = p._image_to_base64(img)
        out_list.append(f"{default_indent}data: {base64_str}\n")

        out_list.append("\n")


def _add_property_if_exists(item: Any, name: str, out_list: List[str], indents: int = 1):
    value = getattr(item, name, None) 
    if value is not None:      
        out_list.append(f"{default_indent * indents}{name}: {_to_single_line_text(value)}\n")


def _to_single_line_text(value: Any) -> str:
    if isinstance(value, str):
        return value.replace("\r\n", "\\r\\n").replace("\n", "\\n")
    else:
        return value

def _add_parent_if_exists(item: DocItem, out_list: List[str]):
    if item.parent is not None:      
        out_list.append(f"{default_indent}parent: {_strip_ref(item.parent.cref)}\n")


def _get_formatting_string(f: Formatting) -> str:
    if f is None:
        return "None"
    
    s = ""
    if f.bold:
        s += "bold "
    if f.italic:
        s += "italic "
    if f.underline:
        s += "underline "
    if f.strikethrough:
        s += "strikethrough "
    if f.script == Script.SUB:
        s += "subscript  "
    if f.script == Script.SUPER:
        s += "superscript  "
        
    return f"[{s.rstrip()}]"


def _strip_ref(ref: str) -> str:
    return ref.lstrip("#/")
    

def verify_documents(doc: DoclingDocument, expected_doc: DoclingDocument, test_filename: str):
    doc_dict = doc.export_to_dict()
    doc_json_string = json.dumps(doc_dict, indent=4)
    doc_json_filename = os.path.join(get_artifacts_dir(), Path(doc.origin.filename).name + ".json")
    with open(doc_json_filename, "w") as f:
        f.write(doc_json_string)   

    doc_tree = get_document_tree(doc) + "\n\n" + get_document_items(doc)
    doc_tree_filename = Path(doc.origin.filename).name + ".aspose"
    doc_filename = os.path.join(get_artifacts_dir(), doc_tree_filename)
    with open(doc_filename, "w", encoding="utf-8") as f:
        f.write(doc_tree)

    if expected_doc is not None:
        expected_doc_tree = get_document_tree(expected_doc) + "\n\n" + get_document_items(expected_doc)
        expected_filename = os.path.join(get_artifacts_dir(), Path(doc.origin.filename).name + ".expected")
        with open(expected_filename, "w", encoding="utf-8") as f:
            f.write(expected_doc_tree)    
    else: 
        expected_filename = os.path.join(os.path.dirname(test_filename), 
                                         Path(doc.origin.filename).name + ".expected")
        with open(expected_filename, "r", encoding="utf-8") as f:
            expected_doc_tree = f.read()    
    
    if (doc_tree != expected_doc_tree):
        d = Differ()
        diff = d.compare(doc_tree.split(), expected_doc_tree.split())
        print("\n".join(diff))

        os.system(f"{DIFF_TOOL} {doc_filename} {expected_filename}")
        pytest.fail("Documents trees are not equal.")


def verify(test_filename: str, compare_with_docling: bool = True):
    aspose_converter = AsposeWordsConverter()
    full_test_filename = get_in_test_dir(test_filename)
    aspose_doc = aspose_converter.convert(full_test_filename)
    if compare_with_docling:
        converter = DocumentConverter()
        expected_doc = converter.convert(full_test_filename).document
    else:
        expected_doc = None
    
    verify_documents(aspose_doc, expected_doc, full_test_filename)
