/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.configdata;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "")
@XmlRootElement(name = "CommonData")
public class CommonData {
    @XmlAttribute(name = "businessType")
    protected BusinessType businessType;

    @XmlAttribute(name = "baseTemplatePath")
    protected String baseTemplatePath;

    @XmlAttribute(name = "iub")
    protected Boolean iub;

    public CommonData() {
    }

    public BusinessType getBusinessType() {
        return this.businessType;
    }

    public void setBusinessType(BusinessType value) {
        this.businessType = value;
    }

    public String getBaseTemplatePath() {
        return this.baseTemplatePath;
    }

    public void setBaseTemplatePath(String value) {
        this.baseTemplatePath = value;
    }

    public Boolean isIub() {
        return this.iub;
    }

    public void setIub(Boolean value) {
        this.iub = value;
    }
}