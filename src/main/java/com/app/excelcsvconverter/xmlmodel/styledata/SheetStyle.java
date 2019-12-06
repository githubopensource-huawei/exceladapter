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
import javax.xml.bind.annotation.XmlSchemaType;
import javax.xml.bind.annotation.XmlType;

@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "",
    propOrder = {"validation", "cellRange", "freezePane", "visible", "rowStyle", "columnStyle", "cellLevelStyle"})
@XmlRootElement(name = "SheetStyle")
public class SheetStyle {
    @XmlElement(name = "Validation", required = true)
    protected List<Validation> validation;

    @XmlElement(name = "CellRange", required = true)
    protected List<CellRange> cellRange;

    @XmlElement(name = "FreezePane", required = true)
    protected FreezePane freezePane;

    @XmlElement(required = true)
    @XmlSchemaType(name = "string")
    protected SheetVisible visible;

    @XmlElement(name = "RowStyle", required = true)
    protected List<RowStyle> rowStyle;

    @XmlElement(name = "ColumnStyle", required = true)
    protected List<ColumnStyle> columnStyle;

    @XmlElement(name = "CellLevelStyle", required = true)
    protected CellLevelStyle cellLevelStyle;

    @XmlAttribute(name = "sheetName", required = true)
    protected String sheetName;

    @XmlAttribute(name = "tabColor")
    protected String tabColor;

    @XmlAttribute(name = "displayGridLines")
    protected Boolean displayGridLines;

    public SheetStyle() {
    }

    public List<Validation> getValidation() {
        if (this.validation == null) {
            this.validation = new ArrayList();
        }

        return this.validation;
    }

    public List<CellRange> getCellRange() {
        if (this.cellRange == null) {
            this.cellRange = new ArrayList();
        }

        return this.cellRange;
    }

    public FreezePane getFreezePane() {
        return this.freezePane;
    }

    public void setFreezePane(FreezePane value) {
        this.freezePane = value;
    }

    public SheetVisible getVisible() {
        return this.visible;
    }

    public void setVisible(SheetVisible value) {
        this.visible = value;
    }

    public List<RowStyle> getRowStyle() {
        if (this.rowStyle == null) {
            this.rowStyle = new ArrayList();
        }

        return this.rowStyle;
    }

    public List<ColumnStyle> getColumnStyle() {
        if (this.columnStyle == null) {
            this.columnStyle = new ArrayList();
        }

        return this.columnStyle;
    }

    public CellLevelStyle getCellLevelStyle() {
        return this.cellLevelStyle;
    }

    public void setCellLevelStyle(CellLevelStyle value) {
        this.cellLevelStyle = value;
    }

    public String getSheetName() {
        return this.sheetName;
    }

    public void setSheetName(String value) {
        this.sheetName = value;
    }

    public String getTabColor() {
        return this.tabColor;
    }

    public void setTabColor(String value) {
        this.tabColor = value;
    }

    public Boolean isDisplayGridLines() {
        return this.displayGridLines;
    }

    public void setDisplayGridLines(Boolean value) {
        this.displayGridLines = value;
    }
}