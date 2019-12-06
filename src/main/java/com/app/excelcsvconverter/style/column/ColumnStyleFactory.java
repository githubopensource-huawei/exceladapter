/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.column;

import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

public class ColumnStyleFactory {

    private List<ColumnStyleOperator> columnStyleOperators = new ArrayList<>();

    private List<ColumnStyleOperator> getColumnStyleOperators() {
        HiddenColumnSetter hiddenColumnSetter = new HiddenColumnSetter();
        columnStyleOperators.add(hiddenColumnSetter);
        ColumnWidthSetter columnWidthSetter = new ColumnWidthSetter();
        columnStyleOperators.add(columnWidthSetter);
        return columnStyleOperators;
    }

    public void setColumnLevelStyle(Sheet sheet, short column, SheetStyle sheetStyle) {
        if (sheetStyle == null) {
            return;
        }
        for (ColumnStyleOperator columnStyleOperator : getColumnStyleOperators()) {
            columnStyleOperator.setStyle(sheet, column, sheetStyle);
        }
    }
}
