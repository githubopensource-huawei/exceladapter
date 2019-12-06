/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.styledata;

import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlAttribute;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "", propOrder = {"sheetStyle", "cellStyleData", "fontData"})
@XmlRootElement(name = "ExcelStyle")
public class ExcelStyle {
    @XmlElement(name = "SheetStyle", required = true)
    protected List<SheetStyle> sheetStyle;

    @XmlElement(name = "CellStyleData", required = true)
    protected List<CellStyleData> cellStyleData;

    @XmlElement(name = "FontData", required = true)
    protected List<FontData> fontData;

    @XmlAttribute(name = "sortedSheetNames")
    protected String sortedSheetNames;

    @XmlAttribute(name = "activeSheet")
    protected String activeSheet;

    @XmlAttribute(name = "defaultFontName")
    protected String defaultFontName;

    @XmlAttribute(name = "defaultFontSize")
    protected Short defaultFontSize;

    public ExcelStyle() {
    }

    public List<SheetStyle> getSheetStyle() {
        if (this.sheetStyle == null) {
            this.sheetStyle = new ArrayList();
        }

        return this.sheetStyle;
    }

    public List<CellStyleData> getCellStyleData() {
        if (this.cellStyleData == null) {
            this.cellStyleData = new ArrayList();
        }

        return this.cellStyleData;
    }

    public List<FontData> getFontData() {
        if (this.fontData == null) {
            this.fontData = new ArrayList();
        }

        return this.fontData;
    }

    public String getSortedSheetNames() {
        return this.sortedSheetNames;
    }

    public void setSortedSheetNames(String value) {
        this.sortedSheetNames = value;
    }

    public String getActiveSheet() {
        return this.activeSheet;
    }

    public void setActiveSheet(String value) {
        this.activeSheet = value;
    }

    public String getDefaultFontName() {
        return this.defaultFontName;
    }

    public void setDefaultFontName(String value) {
        this.defaultFontName = value;
    }

    public Short getDefaultFontSize() {
        return this.defaultFontSize;
    }

    public void setDefaultFontSize(Short value) {
        this.defaultFontSize = value;
    }
}