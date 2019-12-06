/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.row;

import com.app.excelcsvconverter.xmlmodel.styledata.RowStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Row;

import java.util.List;

/**
 * 设置行高
 *
 * @since 2019-07-22
 */
public class RowHeightSetter implements RowStyleOperator {

    @Override
    public void setStyle(Row row, int rowIndex, SheetStyle sheetStyle) {
        List<RowStyle> rowStyles = sheetStyle.getRowStyle();
        for (RowStyle rowStyle : rowStyles) {
            if ((rowStyle.getRowBegin() <= rowIndex && rowStyle.getRowEnd() >= rowIndex) && rowStyle.getHeight() != null) {
                row.setHeight((short) rowStyle.getHeight().intValue());
            }
        }
    }
}