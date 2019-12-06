/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.styledata;

import javax.xml.bind.annotation.XmlEnum;
import javax.xml.bind.annotation.XmlType;

@XmlType(name = "CellStyleType")
@XmlEnum
public enum CellStyleType {
    CUSTOMIZED,
    HEAD_STYLE,
    COLUMN_STYLE,
    DATA_STYLE,
    HYPER_LINK_STYLE,
    COMMENT_STYLE,
    IUB_GROUP_STYLE,
    IUB_COLUMN_STYLE,
    SUMMARY_PATTEN_HYPERLINK,
    SUMMARY_TITLE_HYPERLINK,
    CONFLICT_LINK,
    FEATURE_HYPERLINK,
    SUMMARY_DATA_TOP_LEFT,
    SUMMARY_DATA_CENTER_LEFT;

    CellStyleType() {
    }

    public static CellStyleType fromValue(String v) {
        return valueOf(v);
    }

    public String value() {
        return this.name();
    }
}