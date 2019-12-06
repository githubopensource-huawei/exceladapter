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
@XmlRootElement(name = "CellStyleData")
public class CellStyleData {
    @XmlAttribute(name = "index", required = true)
    protected int index;

    @XmlAttribute(name = "backgroundColor")
    protected String backgroundColor;

    @XmlAttribute(name = "dataFormat")
    protected Short dataFormat;

    @XmlAttribute(name = "dataFormatString")
    protected String dataFormatString;

    @XmlAttribute(name = "alignment")
    protected String alignment;

    @XmlAttribute(name = "verticalAlignment")
    protected String verticalAlignment;

    @XmlAttribute(name = "borderBottom")
    protected String borderBottom;

    @XmlAttribute(name = "borderLeft")
    protected String borderLeft;

    @XmlAttribute(name = "borderTop")
    protected String borderTop;

    @XmlAttribute(name = "borderRight")
    protected String borderRight;

    @XmlAttribute(name = "wrapText")
    protected Boolean wrapText;

    @XmlAttribute(name = "fillPattern")
    protected String fillPattern;

    @XmlAttribute(name = "locked")
    protected Boolean locked;

    public CellStyleData() {
    }

    public int getIndex() {
        return this.index;
    }

    public void setIndex(int value) {
        this.index = value;
    }

    public String getBackgroundColor() {
        return this.backgroundColor;
    }

    public void setBackgroundColor(String value) {
        this.backgroundColor = value;
    }

    public Short getDataFormat() {
        return this.dataFormat;
    }

    public void setDataFormat(Short value) {
        this.dataFormat = value;
    }

    public String getDataFormatString() {
        return this.dataFormatString;
    }

    public void setDataFormatString(String value) {
        this.dataFormatString = value;
    }

    public String getAlignment() {
        return this.alignment;
    }

    public void setAlignment(String value) {
        this.alignment = value;
    }

    public String getVerticalAlignment() {
        return this.verticalAlignment;
    }

    public void setVerticalAlignment(String value) {
        this.verticalAlignment = value;
    }

    public String getBorderBottom() {
        return this.borderBottom;
    }

    public void setBorderBottom(String value) {
        this.borderBottom = value;
    }

    public String getBorderLeft() {
        return this.borderLeft;
    }

    public void setBorderLeft(String value) {
        this.borderLeft = value;
    }

    public String getBorderTop() {
        return this.borderTop;
    }

    public void setBorderTop(String value) {
        this.borderTop = value;
    }

    public String getBorderRight() {
        return this.borderRight;
    }

    public void setBorderRight(String value) {
        this.borderRight = value;
    }

    public Boolean isWrapText() {
        return this.wrapText;
    }

    public void setWrapText(Boolean value) {
        this.wrapText = value;
    }

    public String getFillPattern() {
        return this.fillPattern;
    }

    public void setFillPattern(String value) {
        this.fillPattern = value;
    }

    public Boolean isLocked() {
        return this.locked;
    }

    public void setLocked(Boolean value) {
        this.locked = value;
    }
}