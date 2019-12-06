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
@XmlRootElement(name = "CommentArea")
public class CommentArea {
    @XmlAttribute(name = "dx1")
    protected Integer dx1;

    @XmlAttribute(name = "dx2")
    protected Integer dx2;

    @XmlAttribute(name = "dy1")
    protected Integer dy1;

    @XmlAttribute(name = "dy2")
    protected Integer dy2;

    @XmlAttribute(name = "col1")
    protected Integer col1;

    @XmlAttribute(name = "col2")
    protected Integer col2;

    @XmlAttribute(name = "row1")
    protected Integer row1;

    @XmlAttribute(name = "row2")
    protected Integer row2;

    public CommentArea() {
    }

    public Integer getDx1() {
        return this.dx1;
    }

    public void setDx1(Integer value) {
        this.dx1 = value;
    }

    public Integer getDx2() {
        return this.dx2;
    }

    public void setDx2(Integer value) {
        this.dx2 = value;
    }

    public Integer getDy1() {
        return this.dy1;
    }

    public void setDy1(Integer value) {
        this.dy1 = value;
    }

    public Integer getDy2() {
        return this.dy2;
    }

    public void setDy2(Integer value) {
        this.dy2 = value;
    }

    public Integer getCol1() {
        return this.col1;
    }

    public void setCol1(Integer value) {
        this.col1 = value;
    }

    public Integer getCol2() {
        return this.col2;
    }

    public void setCol2(Integer value) {
        this.col2 = value;
    }

    public Integer getRow1() {
        return this.row1;
    }

    public void setRow1(Integer value) {
        this.row1 = value;
    }

    public Integer getRow2() {
        return this.row2;
    }

    public void setRow2(Integer value) {
        this.row2 = value;
    }
}