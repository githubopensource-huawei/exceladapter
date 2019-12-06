/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.styledata;

import javax.xml.bind.annotation.XmlEnum;
import javax.xml.bind.annotation.XmlType;

@XmlType(name = "SheetVisible")
@XmlEnum
public enum SheetVisible {
    VISIBLE,
    HIDDEN,
    VERY_HIDDEN;

    SheetVisible() {
    }

    public static SheetVisible fromValue(String v) {
        return valueOf(v);
    }

    public String value() {
        return this.name();
    }
}