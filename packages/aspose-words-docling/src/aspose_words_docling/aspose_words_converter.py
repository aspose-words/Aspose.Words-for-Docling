import io
import os
import json
import base64
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any

from PIL import Image, UnidentifiedImageError

from docling_core.types.doc import (
    DoclingDocument, 
    DocumentOrigin,
    Formatting,
    ProvenanceItem,
    DocItemLabel,
    GroupLabel,
    Script,
    ImageRef,
    DocItem,
    TextItem,
    GroupItem    
)

from docling.utils import utils

import aspose.words as aw

from node_util import NodeUtil
from formatting_helper import (is_formatting_equal, get_paragraph_formatting, get_font_formatting,
                               is_title, is_heading, get_header_name, get_heading_level,
                               is_empty_formatting)
from table_converter import TableConverter
from list_converter import ListConverter
from license_manager import LicenseManager


class ParentItem:
    node: DocItem = None
    level: int = None


class Context:
    def __init__(self, parent_nodes: List[ParentItem], list_converter: ListConverter):
        self.parent_nodes.extend(parent_nodes)
        self.list_converter = list_converter
    
    parent_nodes: List[ParentItem] = []
    list_converter: ListConverter


class AsposeWordsConverter:

    def __init__(self):
        self._aspose_doc: aw.Document = None
        self._docling_doc: DoclingDocument = None
        self._parent_nodes: List[ParentItem] = None
        self._list_converter: ListConverter = None


    def convert(self, file: str) -> DoclingDocument:
        license_manager = LicenseManager()
        license_manager.apply_license() 

        self._aspose_doc = aw.Document(file)
        self._aspose_doc.update_list_labels()
        self._docling_doc = self._create_base_document(file)

        self._new_context()
        self._process_document_structure()
        return self._docling_doc


    def _get_context(self) -> Context:
        return Context(self._parent_nodes, self._list_converter)


    def _restore_context(self, context: Context):
        self._parent_nodes = context.parent_nodes
        self._list_converter = context.list_converter


    def _new_context(self):
        self._parent_nodes: List[ParentItem] = []
        self._push_parent_node(self._docling_doc.body, -1)
        self._list_converter = ListConverter(self._docling_doc)


    def _create_base_document(self, source_path: str) -> DoclingDocument:
        # Create document origin
        origin = DocumentOrigin(
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            binary_hash=utils.create_file_hash(Path(source_path)),
            filename=os.path.basename(source_path)
        )
        
        # Create DoclingDocument
        built_in_props = self._aspose_doc.built_in_document_properties
        doc = DoclingDocument(
            name=built_in_props.title or Path(origin.filename).stem or "Untitled Document",
            origin=origin
        )
        
        # Add metadata if available
        """if built_in_props.created_time:
            doc.creation_date = built_in_props.created_time
        if built_in_props.last_saved_time:
            doc.last_modified = built_in_props.last_saved_time TODO"""
        
        return doc
    
    def _process_document_structure(self):
        self._push_parent_node(self._docling_doc.body, -1)

        for sectionNode in self._aspose_doc.sections:
            # Process main body content
            section = sectionNode.as_section()
            self._process_section(section)
            # Process headers and footers
            self._process_headers_footers(section)


    def _process_section(self, section: aw.Section):
        old_context = self._get_context()
        self._new_context()

        if self._aspose_doc.sections.count != 1:
            section_index = self._aspose_doc.sections.index_of(section)
            group = self._docling_doc.add_group(label=GroupLabel.SECTION,
                                                name=f"header-{section_index}",
                                                parent=self._peek_parent_node())
            self._push_parent_node(group, -1)

        self._process_node_collection(section.body.get_child_nodes(aw.NodeType.ANY, False))
        
        self._restore_context(old_context)


    def _process_node_collection(self, nodes: aw.NodeCollection):
        for node in nodes:
            match node.node_type:
                case aw.NodeType.PARAGRAPH:
                    self._process_paragraph(node.as_paragraph())
                case aw.NodeType.TABLE:
                    self._process_table(node.as_table())
                case aw.NodeType.ROW:
                    self._process_node_collection(node.as_row().cells)
                case aw.NodeType.CELL:
                    self._process_node_collection(node.as_cell().get_child_nodes(aw.NodeType.ANY, False))
                case _:
                    print("NOT PROCESSED " + node.node_type) # TODO
                

    def _process_headers_footers(self, section: aw.Section):
        return
        

    def _process_paragraph(self, para: aw.Paragraph):
        self._process_inline_images(para)
        self._process_text_boxes(para)

        while self._list_converter.is_list_finished(para):
            self._list_converter.finish_current_list()
            self._pop_parent_node()

        heading_level = get_heading_level(para)
        if heading_level is not None:
            self._pop_parent_node(heading_level)

        if para.paragraph_format.is_list_item and self._list_converter.is_group_needed(para):
            list_group = self._docling_doc.add_group(label=GroupLabel.LIST,
                                                     name="list",
                                                     parent=self._peek_parent_node())
            self._list_converter.start_list(para)
            self._push_parent_node(list_group)

        is_plain_text, _ = self._get_paragraph_formatting(para)
        is_title_or_heading = is_title(para) or is_heading(para)
        if is_plain_text or is_title_or_heading: 
            para_item = self._process_paragraph_as_plain_text(para)
            if is_title_or_heading:
                self._push_parent_node(para_item, heading_level)
        else:
            self._process_paragraph_as_inline_group(para)


    def _process_inline_images(self, paragraph: aw.Paragraph):
        for item in paragraph.get_child_nodes(aw.NodeType.SHAPE, False):
            shape = item.as_shape()
            if shape.has_image:
                self._process_image(shape)


    def _process_text_boxes(self, paragraph: aw.Paragraph):
        for item in paragraph.get_child_nodes(aw.NodeType.SHAPE, True):
            shape = item.as_shape()
            if shape.shape_type == aw.drawing.ShapeType.TEXT_BOX:
                self._process_textbox(shape)


    def _get_paragraph_formatting(self, para: aw.Paragraph) -> tuple[bool, Formatting]:
        formatting = None
        is_hyperlink = False
        has_runs_before = False
        field_locks = 0
        hard_locks = 0
        nodes_count = 0
        office_math_count = 0

        for item in para.get_child_nodes(aw.NodeType.ANY, False):
            nodes_count += 1
            match item.node_type:
                case aw.NodeType.RUN:
                    if field_locks == 0 and hard_locks == 0:
                        run = item.as_run()
                        if not NodeUtil.is_whitespaces(NodeUtil.get_run_text(run)):
                            if formatting is None:
                                formatting = get_font_formatting(run.font, is_hyperlink)
                            elif not is_formatting_equal(formatting, get_font_formatting(run.font, is_hyperlink)):
                                return (False, None)
                        has_runs_before = True
                case aw.NodeType.FIELD_START:
                    field_locks += 1
                    field = item.as_field_start().get_field()
                    if (field.type == aw.fields.FieldType.FIELD_HYPERLINK and
                        (has_runs_before or field.end != para.last_child)):
                        return (False, None)
                    is_hyperlink = field.type == aw.fields.FieldType.FIELD_HYPERLINK
                case aw.NodeType.FIELD_SEPARATOR:
                     if field_locks > 0:
                        field_locks -= 1
                case aw.NodeType.OFFICE_MATH:
                    office_math_count += 1
                case aw.NodeType.FIELD_END:
                    is_hyperlink = False
                case aw.NodeType.BOOKMARK_START:
                    hard_locks += 1
                case aw.NodeType.BOOKMARK_END:
                     if hard_locks > 0: # TODO: ignore bookmarks
                        hard_locks -= 1
                case (aw.NodeType.SHAPE | aw.NodeType.COMMENT |
                      aw.NodeType.COMMENT_RANGE_START | aw.NodeType.COMMENT_RANGE_END |
                      aw.NodeType.FOOTNOTE):
                    pass
                case _:
                    return (False, None)
        
        is_plain_text = office_math_count == 0 or (office_math_count == 1 and nodes_count == 1)
        if ((is_title(para) or is_heading(para)) and 
            (formatting is not None) and is_empty_formatting(formatting)):
            formatting = None
        return (is_plain_text, formatting)


    def _process_paragraph_as_plain_text(self, para: aw.Paragraph) -> TextItem:
        text = NodeUtil.get_paragraph_text(para)
        label = self._get_paragraph_label(para)
        level = self._get_heading_level(para)
        if text == "":
            formatting = None
        else:
            _, formatting = self._get_paragraph_formatting(para)

        if level != 0:
            item = self._docling_doc.add_heading(level=level,
                                                 parent=self._peek_parent_node(),
                                                 text=text, 
                                                 formatting=formatting)
        else:
            item = self._docling_doc.add_text(label=label,
                                              parent=self._peek_parent_node(),
                                              text=text, 
                                              formatting=formatting)
        
        if (para.paragraph_format.is_list_item and
            para.list_format.list.list_levels[para.list_format.list_level_number].number_style != aw.NumberStyle.BULLET):
            item.marker = para.list_label.label_string
            item.enumerated = True

        return item


    def _process_paragraph_as_inline_group(self, paragraph: aw.Paragraph) -> GroupItem:
        old_parent = self._peek_parent_node()
        if paragraph.paragraph_format.is_list_item:
            list_item = self._docling_doc.add_text(label=DocItemLabel.LIST_ITEM,
                                                   text="",
                                                   parent=self._peek_parent_node())
            self._push_parent_node(list_item)
        inline_group = self._docling_doc.add_inline_group(parent=self._peek_parent_node())
        self._push_parent_node(inline_group)

        prev_text_item: TextItem = None
        locks: int = 0
        hyperlink_field: aw.fields.FieldHyperlink = None
        for item in paragraph.get_child_nodes(aw.NodeType.ANY, False):
            match item.node_type:
                case aw.NodeType.RUN:
                    if locks == 0:
                        prev_text_item = self._process_inline(item, 
                                                              prev_text_item, 
                                                              field_hyperlink=hyperlink_field)
                case aw.NodeType.OFFICE_MATH:
                    if locks == 0:
                        # To process OfficeMath as separate text item set prev_text_item to None.
                        prev_text_item = None 
                        self._process_office_math(item.as_office_math(),
                                                  field_hyperlink=hyperlink_field)
                case aw.NodeType.FIELD_START:
                    locks += 1
                    field_start = item.as_field_start()
                    if (field_start.field_type == aw.fields.FieldType.FIELD_HYPERLINK):
                        self._rstrip_text_item(prev_text_item)
                        prev_text_item = None
                        hyperlink_field = field_start.get_field().as_field_hyperlink()
                case aw.NodeType.FIELD_SEPARATOR:
                    if locks > 0:
                        locks -= 1
                case aw.NodeType.FIELD_END:
                    hyperlink_field = None
                case (aw.NodeType.SHAPE | aw.NodeType.COMMENT | aw.NodeType.COMMENT_RANGE_START | 
                      aw.NodeType.COMMENT_RANGE_END | aw.NodeType.FOOTNOTE):
                    # Common, Footnotes aren't supported by Docling.
                    pass

        self._restore_parent(old_parent)
        self._rstrip_text_item(prev_text_item)

        return inline_group


    def _rstrip_text_item(self, text_item: TextItem):
        if text_item is not None:
            text_item.text = text_item.text.rstrip()        
            text_item.orig = text_item.orig.rstrip()        


    def _process_inline(self, inline: aw.Inline, prev_text_item: TextItem,
                        field_hyperlink: aw.fields.FieldHyperlink = None) -> TextItem:
        if inline.node_type == aw.NodeType.RUN:
            run = inline.as_run()
            inline_text = NodeUtil.get_run_text(run)
            formatting = get_font_formatting(run.font, field_hyperlink is not None)
            ignore_formatting = False
        else:
            inline_text = inline.get_text()
            formatting = Formatting()
            ignore_formatting = True
        
        # Combine runs with identical formatting
        if ((prev_text_item is not None) and 
            (ignore_formatting or is_formatting_equal(prev_text_item.formatting, formatting))):
            prev_text_item.text = prev_text_item.text + inline_text
            prev_text_item.orig = prev_text_item.text
            return prev_text_item

        if prev_text_item:
            prev_text_item.text = prev_text_item.text.rstrip()        
            prev_text_item.orig = prev_text_item.orig.rstrip()        

        if inline_text.strip() == "":
            return prev_text_item
        
        text_item = self._docling_doc.add_text(label=DocItemLabel.PARAGRAPH,
                                               parent=self._peek_parent_node(),
                                               text=inline_text.lstrip(),
                                               formatting=formatting)
        if field_hyperlink:
            text_item.hyperlink = field_hyperlink.address
        
        if field_hyperlink is not None:
            return None
        
        return text_item
    

    def _process_office_math(self, office_math: aw.math.OfficeMath,
                             field_hyperlink: aw.fields.FieldHyperlink = None):
        text = NodeUtil.get_office_math_text(office_math)
        if text.strip() != "":
            text_item = self._docling_doc.add_formula(parent=self._peek_parent_node(),
                                                      text=text.lstrip())
            if field_hyperlink is not None:
                text_item.hyperlink = field_hyperlink.address


    def _get_font_script(self, font: aw.Font) -> Script:
        if font.subscript:
            return Script.SUB
        if font.superscript:
            return Script.SUPER
        return Script.BASELINE
        

    def _process_image(self, shape: aw.drawing.Shape):
        image_data = shape.image_data
        if not image_data or not image_data.image_bytes:
            return
        
        try:
            image_bytes = BytesIO(image_data.image_bytes)
            pil_image = Image.open(image_bytes)
            self._docling_doc.add_picture(parent=self._peek_parent_node(),
                                          image=ImageRef.from_pil(image=pil_image, dpi=72),
                                          caption=None)
        except (UnidentifiedImageError, OSError):
            self._docling_doc.add_picture(parent=self._peek_parent_node(),
                                          caption=None)
     

    def _process_textbox(self, text_box: aw.drawing.Shape):
        assert(text_box.shape_type == aw.drawing.ShapeType.TEXT_BOX)

        old_context = self._get_context()
        self._new_context()

        textbox_group = self._docling_doc.add_group(parent=self._docling_doc.body,
                                                    label=GroupLabel.SECTION,
                                                    name="textbox")
        self._push_parent_node(textbox_group)
        self._process_node_collection(text_box.get_child_nodes(aw.NodeType.ANY, False))

        self._restore_context(old_context)


    def _get_paragraph_label(self, paragraph: aw.Paragraph) -> DocItemLabel:
        match paragraph.paragraph_format.style.style_identifier:
            case (aw.StyleIdentifier.TITLE):
                return DocItemLabel.TITLE
            case (aw.StyleIdentifier.HEADING1 | aw.StyleIdentifier.HEADING2 | aw.StyleIdentifier.HEADING3 |
                  aw.StyleIdentifier.HEADING4 | aw.StyleIdentifier.HEADING5 |
                  aw.StyleIdentifier.HEADING6 | aw.StyleIdentifier.HEADING7 |
                  aw.StyleIdentifier.HEADING8 | aw.StyleIdentifier.HEADING9):
                return DocItemLabel.SECTION_HEADER

        if paragraph.list_format.is_list_item:
            return DocItemLabel.LIST_ITEM
         
        if ((paragraph.first_child is not None) and 
            paragraph.first_child.node_type == aw.NodeType.OFFICE_MATH and
            paragraph.first_child == paragraph.last_child):
            return DocItemLabel.FORMULA

        if paragraph.paragraph_format.style.font.bold:
            return DocItemLabel.SECTION_HEADER
         
        return DocItemLabel.PARAGRAPH

    
    def _get_heading_level(self, paragraph: aw.Paragraph):
        match paragraph.paragraph_format.style.style_identifier:
            case aw.StyleIdentifier.HEADING1:
                return 1
            case aw.StyleIdentifier.HEADING2:
                return 2
            case aw.StyleIdentifier.HEADING3:
                return 3
            case aw.StyleIdentifier.HEADING4:
                return 4
            case aw.StyleIdentifier.HEADING5:
                return 5
            case aw.StyleIdentifier.HEADING6:
                return 6
            case aw.StyleIdentifier.HEADING7:
                return 7
            case aw.StyleIdentifier.HEADING8:
                return 8
            case aw.StyleIdentifier.HEADING9:
                return 9
            
        return 0


    def _process_table(self, table: aw.tables.Table):
        table_converter = TableConverter(self._docling_doc)
        if not table_converter.convert(table, self._peek_parent_node()):
            self._process_node_collection(table.rows)


    def _push_parent_node(self, node: DocItem, level: int = None):
        item = ParentItem()
        item.node = node
        item.level = level
        return self._parent_nodes.append(item)


    def _pop_parent_node(self, heading_level: int = None) -> DocItem:
        if heading_level is None:
            assert(len(self._parent_nodes) > 1)
            return self._parent_nodes.pop().node
        
        item = self._parent_nodes[-1]
        while (item.level is None) or item.level >= heading_level:
            assert(len(self._parent_nodes) > 1)
            self._parent_nodes.pop()
            item = self._parent_nodes[-1]
        return item.node


    def _peek_parent_node(self) -> DocItem:
        return self._parent_nodes[-1].node


    def _restore_parent(self, parent_node: DocItem):
        while self._peek_parent_node() != parent_node:
            node = self._pop_parent_node()        
            if node is None:
                assert(False)
