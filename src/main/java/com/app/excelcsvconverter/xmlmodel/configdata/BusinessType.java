/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.configdata;

import javax.xml.bind.annotation.XmlEnum;
import javax.xml.bind.annotation.XmlType;

@XmlType(name = "BusinessType")
@XmlEnum
public enum BusinessType {
    SUMMARY;

    BusinessType() {
    }

    public static BusinessType fromValue(String v) {
        return valueOf(v);
    }

    public String value() {
        return this.name();
    }
}