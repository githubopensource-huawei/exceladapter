/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.configdata;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "", propOrder = {"commonData"})
@XmlRootElement(name = "ConfigData")
public class ConfigData {
    @XmlElement(name = "CommonData", required = true)
    protected CommonData commonData;

    public ConfigData() {
    }

    public CommonData getCommonData() {
        return this.commonData;
    }

    public void setCommonData(CommonData value) {
        this.commonData = value;
    }
}