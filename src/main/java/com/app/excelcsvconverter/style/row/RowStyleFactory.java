/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.row;

import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.List;

/**
 * 创建行级样式的工厂
 *
 * @since 2019-07-22
 */
public class RowStyleFactory {

    private List<RowStyleOperator> rowStyleOperators = new ArrayList<>();

    private List<RowStyleOperator> getRowStyleOperators() {
        RowHeightSetter rowHeightSetter = new RowHeightSetter();
        rowStyleOperators.add(rowHeightSetter);
        HiddenRowSetter hiddenRowSetter = new HiddenRowSetter();
        rowStyleOperators.add(hiddenRowSetter);
        return rowStyleOperators;
    }

    public void setRowLevelStyle(Row row, int rowIndex, SheetStyle sheetStyle) {
        if (sheetStyle == null) {
            return;
        }
        for (RowStyleOperator rowStyleOperator : getRowStyleOperators()) {
            rowStyleOperator.setStyle(row, rowIndex, sheetStyle);
        }
    }
}