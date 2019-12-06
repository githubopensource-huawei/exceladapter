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
@XmlRootElement(name = "FreezePane")
public class FreezePane {
    @XmlAttribute(name = "colNum", required = true)
    protected int colNum;

    @XmlAttribute(name = "rowNum", required = true)
    protected int rowNum;

    @XmlAttribute(name = "firstColNum", required = true)
    protected int firstColNum;

    @XmlAttribute(name = "firstRowNum", required = true)
    protected int firstRowNum;

    public FreezePane() {
    }

    public int getColNum() {
        return this.colNum;
    }

    public void setColNum(int value) {
        this.colNum = value;
    }

    public int getRowNum() {
        return this.rowNum;
    }

    public void setRowNum(int value) {
        this.rowNum = value;
    }

    public int getFirstColNum() {
        return this.firstColNum;
    }

    public void setFirstColNum(int value) {
        this.firstColNum = value;
    }

    public int getFirstRowNum() {
        return this.firstRowNum;
    }

    public void setFirstRowNum(int value) {
        this.firstRowNum = value;
    }
}