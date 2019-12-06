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
@XmlType(name = "", propOrder = {"commentArea"})
@XmlRootElement(name = "CommentData")
public class CommentData {
    @XmlElement(name = "CommentArea", required = true)
    protected CommentArea commentArea;

    @XmlAttribute(name = "row", required = true)
    protected int row;

    @XmlAttribute(name = "column", required = true)
    protected int column;

    @XmlAttribute(name = "comment", required = true)
    protected String comment;

    @XmlAttribute(name = "fontIndex", required = true)
    protected int fontIndex;

    @XmlAttribute(name = "cellStyleType")
    protected CellStyleType cellStyleType;

    public CommentData() {
    }

    public CommentArea getCommentArea() {
        return this.commentArea;
    }

    public void setCommentArea(CommentArea value) {
        this.commentArea = value;
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

    public String getComment() {
        return this.comment;
    }

    public void setComment(String value) {
        this.comment = value;
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