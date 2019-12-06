package com.app.excelcsvconverter.style.cell;

import com.app.excelcsvconverter.style.CellStyleCreator;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 单元格级别的样式
 *
 * @since 2019-07-22
 */
public interface CellStyleOperator {

    /**
     * Set the cell level style
     *
     * @param cell         the cell
     * @param rowIndex     the row index
     * @param colIndex     the column index
     * @param sheetStyle   the sheetStyle
     * @param cellStyleMap the cellStyleMap
     * @param styleCreator the styleCreator
     */
    void setStyle(Cell cell, int rowIndex, short colIndex, SheetStyle sheetStyle, Map<String, CellStyle> cellStyleMap,
        CellStyleCreator styleCreator);
}