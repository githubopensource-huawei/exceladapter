/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.sheet;

import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * 页签级别的样式
 *
 * @since 2019-07-19
 */
public interface SheetStyleOperator {

    /**
     * Set the sheet level style
     *
     * @param sheet
     * @param sheetStyle
     */
    void setStyle(Sheet sheet, SheetStyle sheetStyle);
}