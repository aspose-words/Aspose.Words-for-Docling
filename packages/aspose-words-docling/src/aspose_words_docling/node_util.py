import string
import aspose.words as aw

class NodeUtil:
    # Returns true if element is block-level type.
    def is_block_level_node(node: aw.Node) -> bool:
        match node.node_type:
            case (aw.NodeType.PARAGRAPH | aw.NodeType.TABLE | aw.NodeType.BOOKMARK_START |
                  aw.NodeType.BOOKMARK_END | aw.NodeType.COMMENT_RANGE_START | aw.NodeType.COMMENT_RANGE_END |
                  aw.NodeType.MOVE_FROM_RANGE_START | aw.NodeType.MOVE_FROM_RANGE_END | 
                  aw.NodeType.MOVE_TO_RANGE_START | aw.NodeType.MOVE_TO_RANGE_END |
                  aw.NodeType.EDITABLE_RANGE_START | aw.NodeType.EDITABLE_RANGE_END):
                return True
            case aw.NodeType.STRUCTURED_DOCUMENT_TAG:
                return node.as_structured_document_tag().level == aw.Markup.MarkupLevel.Block
            case aw.NodeType.SYSTEM:
                return isinstance(node, aw.Markup.SdtMarkerStart) | isinstance(node, aw.Markup.SdtMarkerEnd)
            case _:
                return False
            

    def get_paragraph_text(para: aw.Paragraph) -> str:
        return NodeUtil._get_composite_node_text(para)


    def _get_composite_node_text(node: aw.CompositeNode) -> str:
        text = ""
        locks: int = 0
        for item in node.get_child_nodes(aw.NodeType.ANY, False):
            match item.node_type:
                case aw.NodeType.RUN:
                    if locks == 0:
                        text += NodeUtil.get_run_text(item.as_run())
                case aw.NodeType.FIELD_START:
                    locks += 1
                case aw.NodeType.FIELD_SEPARATOR:
                    if locks > 0:
                        locks -= 1
                case aw.NodeType.OFFICE_MATH:
                    if locks == 0:
                        text += NodeUtil.get_office_math_text(item.as_office_math())
                case aw.NodeType.PARAGRAPH:
                    if locks == 0:
                        text += NodeUtil._get_composite_node_text(item.as_paragraph())
                case (aw.NodeType.COMMENT | aw.NodeType.FOOTNOTE):
                    # Not supported by Docling.
                    pass

        stripped_text = text.strip()
        if stripped_text == "":
            return ""
        else:
            return text 


    def get_run_text(run: aw.Run) -> str:
        soft_line_break = '\v'
        return run.text.replace(soft_line_break, '\n')


    def get_office_math_text(om: aw.math.OfficeMath) -> str:
        mso = aw.saving.MarkdownSaveOptions()
        mso.office_math_export_mode = aw.saving.MarkdownOfficeMathExportMode.LATEX
        return om.to_string(mso) # TODO https://issue.auckland.dynabic.com/issues/WORDSNET-28695


    def get_cell_text(cell: aw.tables.Cell) -> str:
        text = ""
        for para in cell.paragraphs:
            text += NodeUtil.get_paragraph_text(para.as_paragraph())
            if para != cell.last_paragraph:
                text += "\n"

        return text
    

    def is_whitespaces(s: str) -> bool:
        return all(char in string.whitespace for char in s)
