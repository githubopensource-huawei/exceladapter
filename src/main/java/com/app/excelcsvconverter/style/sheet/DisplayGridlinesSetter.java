/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.sheet;

import com.app.excelcsvconverter.util.ExcelUtil;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * 设置单元格框是否可见
 *
 * @since 2019-07-22
 */
public class DisplayGridlinesSetter implements SheetStyleOperator {

    @Override
    public void setStyle(Sheet sheet, SheetStyle sheetStyle) {
        if (sheetStyle.isDisplayGridLines() == null) {
            return;
        }
        if (ExcelUtil.isCoverSheet(sheet.getSheetName())) {
            sheet.setDisplayGridlines(false);
        } else {
            sheet.setDisplayGridlines(sheetStyle.isDisplayGridLines());
        }
    }
}