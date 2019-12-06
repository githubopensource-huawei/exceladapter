/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.sheet;

import com.app.excelcsvconverter.xmlmodel.styledata.FreezePane;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * 设置冻结的区域
 *
 * @since 2019-07-22
 */
public class FreezePaneSetter implements SheetStyleOperator {

    @Override
    public void setStyle(Sheet sheet, SheetStyle sheetStyle) {
        FreezePane pane = sheetStyle.getFreezePane();
        if (pane == null) {
            return;
        }
        sheet.createFreezePane(pane.getColNum(), pane.getRowNum(), pane.getFirstColNum(), pane.getFirstRowNum());
    }
}