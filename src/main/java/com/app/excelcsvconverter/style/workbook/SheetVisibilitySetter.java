/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.workbook;

import com.app.excelcsvconverter.xmlmodel.styledata.ExcelStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;
import com.google.common.base.Strings;

import java.util.List;

import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 设置页签的可见性
 *
 * @since 2019-07-22
 */
public class SheetVisibilitySetter implements WorkbookStyleOperator {

    @Override
    public void setStyle(Workbook workbook, ExcelStyle excelStyle) {
        List<SheetStyle> sheetStyles = excelStyle.getSheetStyle();
        for (SheetStyle sheetStyle : sheetStyles) {
            String sheetName = sheetStyle.getSheetName();
            int sheetIndex = workbook.getSheetIndex(sheetName);
            if (Strings.isNullOrEmpty(sheetName) || sheetStyle.getVisible() == null || sheetIndex == -1) {
                continue;
            }
            SheetVisibility[] sheetVisibilities = SheetVisibility.values();
            for (SheetVisibility sheetVisibility : sheetVisibilities) {
                if (sheetStyle.getVisible().value().equals(sheetVisibility.name())) {
                    workbook.setSheetVisibility(sheetIndex, sheetVisibility);
                }
            }
        }
        // FileIdentification页签特殊处理，设为深度隐藏页签
        int sheetIndex = workbook.getSheetIndex("FileIdentification");
        if (sheetIndex != -1) {
            workbook.setSheetVisibility(sheetIndex, SheetVisibility.VERY_HIDDEN);
        }
    }
}
