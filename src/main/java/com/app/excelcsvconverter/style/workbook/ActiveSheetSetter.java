/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.workbook;

import com.app.excelcsvconverter.xmlmodel.styledata.ExcelStyle;
import com.google.common.base.Strings;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * 设置首先显示的页签
 *
 * @since 2019-07-22
 */
public class ActiveSheetSetter implements WorkbookStyleOperator {

    @Override
    public void setStyle(Workbook workbook, ExcelStyle excelStyle) {
        String sheetName = excelStyle.getActiveSheet();
        if (Strings.isNullOrEmpty(sheetName)) {
            return;
        }
        int index = workbook.getSheetIndex(sheetName);
        if (index == -1) {
            return;
        }
        workbook.setActiveSheet(index);
        workbook.setSelectedTab(index);
    }
}
