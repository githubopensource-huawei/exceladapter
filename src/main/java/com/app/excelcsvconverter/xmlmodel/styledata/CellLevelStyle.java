/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.styledata;

import java.util.ArrayList;
import java.util.List;

import javax.xml.bind.annotation.XmlAccessType;
import javax.xml.bind.annotation.XmlAccessorType;
import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import javax.xml.bind.annotation.XmlType;

@XmlAccessorType(XmlAccessType.FIELD)
@XmlType(name = "", propOrder = {"styleData", "commentData", "hyperlinkData"})
@XmlRootElement(name = "CellLevelStyle")
public class CellLevelStyle {
    @XmlElement(name = "StyleData", required = true)
    protected List<StyleData> styleData;

    @XmlElement(name = "CommentData", required = true)
    protected List<CommentData> commentData;

    @XmlElement(name = "HyperlinkData", required = true)
    protected List<HyperlinkData> hyperlinkData;

    public CellLevelStyle() {
    }

    public List<StyleData> getStyleData() {
        if (this.styleData == null) {
            this.styleData = new ArrayList();
        }

        return this.styleData;
    }

    public List<CommentData> getCommentData() {
        if (this.commentData == null) {
            this.commentData = new ArrayList();
        }

        return this.commentData;
    }

    public List<HyperlinkData> getHyperlinkData() {
        if (this.hyperlinkData == null) {
            this.hyperlinkData = new ArrayList();
        }

        return this.hyperlinkData;
    }
}