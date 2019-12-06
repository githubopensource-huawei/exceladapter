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
@XmlRootElement(name = "ColumnStyle")
public class ColumnStyle {
    @XmlAttribute(name = "colBegin", required = true)
    protected int colBegin;

    @XmlAttribute(name = "colEnd", required = true)
    protected int colEnd;

    @XmlAttribute(name = "width")
    protected Integer width;

    @XmlAttribute(name = "hidden")
    protected Boolean hidden;

    public ColumnStyle() {
    }

    public int getColBegin() {
        return this.colBegin;
    }

    public void setColBegin(int value) {
        this.colBegin = value;
    }

    public int getColEnd() {
        return this.colEnd;
    }

    public void setColEnd(int value) {
        this.colEnd = value;
    }

    public Integer getWidth() {
        return this.width;
    }

    public void setWidth(Integer value) {
        this.width = value;
    }

    public Boolean isHidden() {
        return this.hidden;
    }

    public void setHidden(Boolean value) {
        this.hidden = value;
    }
}