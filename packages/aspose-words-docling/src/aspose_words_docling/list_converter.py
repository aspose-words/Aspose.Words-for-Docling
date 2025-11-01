from typing import List
import aspose.words as aw
from node_util import NodeUtil

from docling_core.types.doc import (
    DoclingDocument,
    GroupItem   
)


class ListInfo:
    def __init__(self, para: aw.Paragraph):
        self.paragraph: aw.Paragraph = para
        self.list_id: int = para.list_format.list.list_id
        self.list_level: aw.lists.ListLevel = para.list_format.list_level
        self.list_level_number: int = para.list_format.list_level_number

    
class ListConverter:
    def __init__(self, docling_doc: DoclingDocument):
        self._docling_doc: DoclingDocument = docling_doc
        self._list_items: List[ListInfo] = []


    def is_group_needed(self, para: aw.Paragraph) -> bool:
        if not para.is_list_item:
            assert(False)
        
        current_list_info = self._peek_list()
        list_info = ListInfo(para)
        return ((current_list_info is None) or
                current_list_info.list_id != list_info.list_id or
                current_list_info.list_level_number < list_info.list_level_number)


    def is_list_finished(self, next_para: aw.Paragraph) -> bool:
        current_list_info = self._peek_list()
        if current_list_info is None:
            return False
        if not next_para.is_list_item:
            return True 

        list_info = ListInfo(next_para)
        return (current_list_info.list_id != list_info.list_id or
                current_list_info.list_level_number > list_info.list_level_number)


    @property
    def current_list_id(self) -> int:
        list_info = self._peek_list()
        if list_info is not None:
            return list_info.list_id
        return None


    def start_list(self, para: aw.Paragraph):
        self._push_list(para)


    def finish_current_list(self):
        self._pop_list()


    def finish_all_lists(self):
        self._list_items.clear()


    def list_group_started(self, group: GroupItem, para: aw.Paragraph):
        self._push_list(para)
        return


    def list_group_finished(self, group: GroupItem, para: aw.Paragraph):
        self._pop_list(para)
        return


    def _push_list(self, paragraph: aw.Paragraph):
        self._list_items.append(ListInfo(paragraph))


    def _pop_list(self):
        assert(len(self._list_items) != 0)
        self._list_items.pop()


    def _peek_list(self) -> ListInfo:
        if len(self._list_items) != 0:
            return self._list_items[-1]
        else:
            return None
