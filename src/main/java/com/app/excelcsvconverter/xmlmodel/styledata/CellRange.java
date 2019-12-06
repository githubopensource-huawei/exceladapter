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
@XmlRootElement(name = "CellRange")
public class CellRange {
    @XmlAttribute(name = "rowBegin", required = true)
    protected int rowBegin;

    @XmlAttribute(name = "rowEnd", required = true)
    protected int rowEnd;

    @XmlAttribute(name = "colBegin", required = true)
    protected int colBegin;

    @XmlAttribute(name = "colEnd", required = true)
    protected int colEnd;

    public CellRange() {
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
}