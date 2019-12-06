/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.parser;

/**
 * The enum Excel format.
 *
 * @since 2019-08-21
 */
public enum ExcelFormat {
    /**
     * Only support reading
     */
    OLE2,
    /**
     * Support reading and writing
     */
    OOXML,
    /**
     * Na excel format.
     */
    NA
}
