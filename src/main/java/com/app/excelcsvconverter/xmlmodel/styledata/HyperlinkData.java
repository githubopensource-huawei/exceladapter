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
@XmlRootElement(name = "HyperlinkData")
public class HyperlinkData {
    @XmlAttribute(name = "cellStyleIndex", required = true)
    protected int cellStyleIndex;

    @XmlAttribute(name = "fontIndex", required = true)
    protected int fontIndex;

    @XmlAttribute(name = "row", required = true)
    protected int row;

    @XmlAttribute(name = "column", required = true)
    protected int column;

    @XmlAttribute(name = "linkSheetName", required = true)
    protected String linkSheetName;

    @XmlAttribute(name = "linkRow", required = true)
    protected int linkRow;

    @XmlAttribute(name = "linkColumn", required = true)
    protected int linkColumn;

    @XmlAttribute(name = "cellStyleType")
    protected CellStyleType cellStyleType;

    public HyperlinkData() {
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

    public int getRow() {
        return this.row;
    }

    public void setRow(int value) {
        this.row = value;
    }

    public int getColumn() {
        return this.column;
    }

    public void setColumn(int value) {
        this.column = value;
    }

    public String getLinkSheetName() {
        return this.linkSheetName;
    }

    public void setLinkSheetName(String value) {
        this.linkSheetName = value;
    }

    public int getLinkRow() {
        return this.linkRow;
    }

    public void setLinkRow(int value) {
        this.linkRow = value;
    }

    public int getLinkColumn() {
        return this.linkColumn;
    }

    public void setLinkColumn(int value) {
        this.linkColumn = value;
    }

    public CellStyleType getCellStyleType() {
        return this.cellStyleType == null ? CellStyleType.CUSTOMIZED : this.cellStyleType;
    }

    public void setCellStyleType(CellStyleType value) {
        this.cellStyleType = value;
    }
}