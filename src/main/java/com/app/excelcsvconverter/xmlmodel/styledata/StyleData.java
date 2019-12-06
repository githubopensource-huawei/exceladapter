/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.styledata;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "", propOrder = {"cellRange"})
@XmlRootElement(name = "StyleData")
public class StyleData {
    @XmlElement(name = "CellRange", required = true)
    protected CellRange cellRange;

    @XmlAttribute(name = "cellStyleIndex", required = true)
    protected int cellStyleIndex;

    @XmlAttribute(name = "fontIndex", required = true)
    protected int fontIndex;

    @XmlAttribute(name = "cellStyleType")
    protected CellStyleType cellStyleType;

    public StyleData() {
    }

    public CellRange getCellRange() {
        return this.cellRange;
    }

    public void setCellRange(CellRange value) {
        this.cellRange = value;
    }

    public int getCellStyleIndex() {
        return this.cellStyleIndex;
    }

    public void setCellStyleIndex(int value) {
        this.cellStyleIndex = value;
    }

    public int getFontIndex() {
        return this.fontIndex;
    }

    public void setFontIndex(int value) {
        this.fontIndex = value;
    }

    public CellStyleType getCellStyleType() {
        return this.cellStyleType == null ? CellStyleType.CUSTOMIZED : this.cellStyleType;
    }

    public void setCellStyleType(CellStyleType value) {
        this.cellStyleType = value;
    }
}