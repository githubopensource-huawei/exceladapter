/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.column;

import com.app.excelcsvconverter.xmlmodel.styledata.ColumnStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

/**
 * 设置列是否隐藏
 *
 * @since 2019-07-22
 */
public class HiddenColumnSetter implements ColumnStyleOperator {

    @Override
    public void setStyle(Sheet sheet, short colIndex, SheetStyle sheetStyle) {
        List<ColumnStyle> columnStyles = sheetStyle.getColumnStyle();
        for (ColumnStyle columnStyle : columnStyles) {
            if ((colIndex >= columnStyle.getColBegin() && colIndex <= columnStyle.getColEnd())
                && columnStyle.isHidden() != null) {
                sheet.setColumnHidden(colIndex, columnStyle.isHidden());
            }
        }
    }
}
