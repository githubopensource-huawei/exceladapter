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
@XmlRootElement(name = "Validation")
public class Validation {
    @XmlElement(name = "CellRange", required = true)
    protected CellRange cellRange;

    @XmlAttribute(name = "validationType", required = true)
    protected int validationType;

    @XmlAttribute(name = "operator", required = true)
    protected int operator;

    @XmlAttribute(name = "attrs", required = true)
    protected String attrs;

    @XmlAttribute(name = "attrs2", required = true)
    protected String attrs2;

    @XmlAttribute(name = "promptBoxTitle")
    protected String promptBoxTitle;

    @XmlAttribute(name = "promptBoxText")
    protected String promptBoxText;

    @XmlAttribute(name = "errorBoxTitle")
    protected String errorBoxTitle;

    @XmlAttribute(name = "errorBoxText")
    protected String errorBoxText;

    public Validation() {
    }

    public CellRange getCellRange() {
        return this.cellRange;
    }

    public void setCellRange(CellRange value) {
        this.cellRange = value;
    }

    public int getValidationType() {
        return this.validationType;
    }

    public void setValidationType(int value) {
        this.validationType = value;
    }

    public int getOperator() {
        return this.operator;
    }

    public void setOperator(int value) {
        this.operator = value;
    }

    public String getAttrs() {
        return this.attrs;
    }

    public void setAttrs(String value) {
        this.attrs = value;
    }

    public String getAttrs2() {
        return this.attrs2;
    }

    public void setAttrs2(String value) {
        this.attrs2 = value;
    }

    public String getPromptBoxTitle() {
        return this.promptBoxTitle;
    }

    public void setPromptBoxTitle(String value) {
        this.promptBoxTitle = value;
    }

    public String getPromptBoxText() {
        return this.promptBoxText;
    }

    public void setPromptBoxText(String value) {
        this.promptBoxText = value;
    }

    public String getErrorBoxTitle() {
        return this.errorBoxTitle;
    }

    public void setErrorBoxTitle(String value) {
        this.errorBoxTitle = value;
    }

    public String getErrorBoxText() {
        return this.errorBoxText;
    }

    public void setErrorBoxText(String value) {
        this.errorBoxText = value;
    }
}