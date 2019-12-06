/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.row;

import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Row;

/**
 * 行级别的样式
 *
 * @since 2019-07-19
 */
public interface RowStyleOperator {

    /**
     * Set the row level style
     *
     * @param row        the row
     * @param rowIndex   the rowIndex
     * @param sheetStyle the sheetStyle
     */
    void setStyle(Row row, int rowIndex, SheetStyle sheetStyle);
}