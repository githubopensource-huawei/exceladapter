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
@XmlRootElement(name = "FontData")
public class FontData {
    @XmlAttribute(name = "index", required = true)
    protected int index;

    @XmlAttribute(name = "fontName")
    protected String fontName;

    @XmlAttribute(name = "fontSize")
    protected Short fontSize;

    @XmlAttribute(name = "color")
    protected String color;

    @XmlAttribute(name = "bold")
    protected Boolean bold;

    @XmlAttribute(name = "underline")
    protected Byte underline;

    public FontData() {
    }

    public int getIndex() {
        return this.index;
    }

    public void setIndex(int value) {
        this.index = value;
    }

    public String getFontName() {
        return this.fontName;
    }

    public void setFontName(String value) {
        this.fontName = value;
    }

    public Short getFontSize() {
        return this.fontSize;
    }

    public void setFontSize(Short value) {
        this.fontSize = value;
    }

    public String getColor() {
        return this.color;
    }

    public void setColor(String value) {
        this.color = value;
    }

    public Boolean isBold() {
        return this.bold;
    }

    public void setBold(Boolean value) {
        this.bold = value;
    }

    public Byte getUnderline() {
        return this.underline;
    }

    public void setUnderline(Byte value) {
        this.underline = value;
    }
}