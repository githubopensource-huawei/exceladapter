/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.column;

import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Sheet;

public interface ColumnStyleOperator {

    void setStyle(Sheet sheet, short colIndex, SheetStyle sheetStyle);
}
