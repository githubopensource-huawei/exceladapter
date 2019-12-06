/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.styledata;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "")
@XmlRootElement(name = "RowStyle")
public class RowStyle {
    @XmlAttribute(name = "rowBegin", required = true)
    protected int rowBegin;

    @XmlAttribute(name = "rowEnd", required = true)
    protected int rowEnd;

    @XmlAttribute(name = "height")
    protected Integer height;

    @XmlAttribute(name = "hidden")
    protected Boolean hidden;

    public RowStyle() {
    }

    public int getRowBegin() {
        return this.rowBegin;
    }

    public void setRowBegin(int value) {
        this.rowBegin = value;
    }

    public int getRowEnd() {
        return this.rowEnd;
    }

    public void setRowEnd(int value) {
        this.rowEnd = value;
    }

    public Integer getHeight() {
        return this.height;
    }

    public void setHeight(Integer value) {
        this.height = value;
    }

    public Boolean isHidden() {
        return this.hidden;
    }

    public void setHidden(Boolean value) {
        this.hidden = value;
    }
}