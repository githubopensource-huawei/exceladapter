package com.app.excelcsvconverter.style.cell;

import com.app.excelcsvconverter.style.CellStyleCreator;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 创建单元格级样式的工厂
 *
 * @since 2019-07-22
 */
public class CellStyleFactory {

    private List<CellStyleOperator> cellStyleOperators = new ArrayList<>();

    private List<CellStyleOperator> getCellStyleOperators() {
        HyperlinkSetter hyperlinkSetter = new HyperlinkSetter();
        cellStyleOperators.add(hyperlinkSetter);
        return cellStyleOperators;
    }

    public void setCellLevelStyle(Cell cell, int rowIndex, short column, SheetStyle sheetStyle,
        Map<String, CellStyle> cellStyleMap, CellStyleCreator styleCreator) {
        if (sheetStyle == null) {
            return;
        }
        for (CellStyleOperator cellStyleOperator : getCellStyleOperators()) {
            cellStyleOperator.setStyle(cell, rowIndex, column, sheetStyle, cellStyleMap, styleCreator);
        }
    }
}