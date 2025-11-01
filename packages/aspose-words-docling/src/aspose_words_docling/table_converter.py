import io
import os
import json
import base64
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Optional, Tuple, Any

import aspose.words as aw
from node_util import NodeUtil

from docling_core.types.doc import (
    DoclingDocument,
    TableData, 
    TableCell,
    TextItem    
)

class TableConverter:
    def __init__(self, docling_doc: DoclingDocument):
        self._docling_doc = docling_doc
        return
    
    def convert(self, table: aw.tables.Table, parent: Any) -> bool:
        row_index = 0
        num_cols = 0
        table_cells: List[TableCell] = []
        table.convert_to_horizontally_merged_cells()

        for row_item in table.rows:
            row = row_item.as_row()
            col_index = 0
            for cell_item in row.cells:
                cell = cell_item.as_cell()
                cell_text = NodeUtil.get_cell_text(cell)
                if ((cell.cell_format.horizontal_merge in [aw.tables.CellMerge.NONE, aw.tables.CellMerge.FIRST]) and
                    (cell.cell_format.vertical_merge in [aw.tables.CellMerge.NONE, aw.tables.CellMerge.FIRST])):
                    col_span = self._get_col_span(row, col_index)
                    row_span = self._get_row_span(table.rows, row_index, col_index)
                    table_cell = TableCell(start_col_offset_idx=col_index,
                                        end_col_offset_idx=col_index + col_span,
                                        start_row_offset_idx=row_index,
                                        end_row_offset_idx=row_index + row_span,
                                        text=cell_text,
                                        column_header=row_index == 0,
                                        col_span=col_span,
                                        row_span=row_span)
                    table_cells.append(table_cell)
                col_index += 1

            num_cols = max(num_cols, col_index)
            row_index += 1

        # Do like Docling - ignore tables with one cell. 
        if len(table_cells) == 1:
            return False

        if len(table_cells) != 0:
            table_data = TableData(table_cells=table_cells,
                                   num_rows= row_index,
                                   num_cols=num_cols)
            self._docling_doc.add_table(data=table_data, parent=parent)

        return True


    def _get_col_span(self, row: aw.tables.Row, cell_index: int) -> int:
        span = 1
        if row.cells[cell_index].cell_format.horizontal_merge == aw.tables.CellMerge.FIRST:
            for i in range(cell_index + 1, row.cells.count):
                if row.cells[i].cell_format.horizontal_merge == aw.tables.CellMerge.PREVIOUS:
                    span += 1
                else:
                    break
            
        return span


    def _get_row_span(self, rows: aw.tables.RowCollection, row_index: int, col_index: int) -> int:
        span = 1
        if rows[row_index].cells[col_index].cell_format.vertical_merge == aw.tables.CellMerge.FIRST:
            for i in range(row_index + 1, rows.count):
                if rows[i].cells[col_index].cell_format.vertical_merge == aw.tables.CellMerge.PREVIOUS:
                    span += 1
                else:
                    break
            
        return span
