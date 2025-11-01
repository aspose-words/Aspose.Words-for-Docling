from docling_core.types.doc import Formatting, Script
import aspose.words as aw


def is_formatting_equal(formatting1: Formatting, formatting2: Formatting, 
                        none_as_empty: bool = False) -> bool:
    if (formatting1 is None) and (formatting2 is None):
        return True

    if (formatting1 is None) != (formatting2 is None) and not none_as_empty:
        return False

    if formatting1 is None:
        formatting1 = Formatting()
    if formatting2 is None:
        formatting2 = Formatting()

    return  (formatting1.bold == formatting2.bold and
            formatting1.italic == formatting2.italic and
            formatting1.underline == formatting2.underline)
            # NOTE: Docling docx backend ignores strikethrough and scripts, so do the same. 
            # formatting1.strikethrough == formatting2.strikethrough and 
            # formatting1.script == formatting2.script)


def is_heading(para: aw.Paragraph) -> bool:
    match para.paragraph_format.style.style_identifier:
        case (aw.StyleIdentifier.HEADING1 | aw.StyleIdentifier.HEADING2 | aw.StyleIdentifier.HEADING3 |
              aw.StyleIdentifier.HEADING4 | aw.StyleIdentifier.HEADING5 |
              aw.StyleIdentifier.HEADING6 | aw.StyleIdentifier.HEADING7 |
              aw.StyleIdentifier.HEADING8 | aw.StyleIdentifier.HEADING9):
            return True
        case _:
            return False


def get_heading_level(para: aw.Paragraph) -> int:
    match para.paragraph_format.style.style_identifier:
        case aw.StyleIdentifier.TITLE:
            return 0
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
        case _:
            return None


def get_header_name(para: aw.Paragraph) -> str:
    level = get_heading_level(para)
    if level is None:
        return None
    return f"header-{level}"


def is_title(para: aw.Paragraph) -> bool:
    return para.paragraph_format.style.style_identifier == aw.StyleIdentifier.TITLE


def is_heading1(para: aw.Paragraph) -> bool:
    return para.paragraph_format.style.style_identifier == aw.StyleIdentifier.HEADING1


def get_paragraph_formatting(para: aw.Paragraph) -> Formatting:
    assert(False) # TODO
    if is_title(para) or is_heading(para):
        return None
    return Formatting()            


def get_font_formatting(font: aw.Font, 
                        is_hyperlink: bool = False, 
                        none_for_empty: bool = False) -> Formatting:
    formatting = Formatting(bold=font.bold,
                            italic=font.italic,
                            underline=not is_hyperlink and font.underline != aw.Underline.NONE)
                            # NOTE: Docling docx backend ignores strikethrough and scripts, so do the same. 
                            # strikethrough=font.strike_through,
                            # script=self._get_font_script(font))
    if none_for_empty and is_empty_formatting(formatting):
        formatting = None
    return formatting


def is_empty_formatting(formatting: Formatting) -> bool:
    return not (formatting.bold or formatting.italic or formatting.underline or
                formatting.strikethrough or formatting.script != Script.BASELINE)
