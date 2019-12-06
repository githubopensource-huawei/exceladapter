/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.workbook;

import com.app.excelcsvconverter.xmlmodel.styledata.ExcelStyle;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * workbook级别的样式
 *
 * @since 2019-07-22
 */
public interface WorkbookStyleOperator {

    /**
     * Set the workbook level style
     *
     * @param workbook
     * @param excelStyle
     */
    void setStyle(Workbook workbook, ExcelStyle excelStyle);
}
