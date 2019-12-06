/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.styledata;

import javax.xml.bind.annotation.XmlRegistry;

@XmlRegistry
public class ObjectFactory {
    public ObjectFactory() {
    }

    public Validation createValidation() {
        return new Validation();
    }

    public CellRange createCellRange() {
        return new CellRange();
    }

    public FreezePane createFreezePane() {
        return new FreezePane();
    }

    public ColumnStyle createColumnStyle() {
        return new ColumnStyle();
    }

    public CellLevelStyle createCellLevelStyle() {
        return new CellLevelStyle();
    }

    public StyleData createStyleData() {
        return new StyleData();
    }

    public CellStyleData createCellStyleData() {
        return new CellStyleData();
    }

    public FontData createFontData() {
        return new FontData();
    }

    public CommentData createCommentData() {
        return new CommentData();
    }

    public HyperlinkData createHyperlinkData() {
        return new HyperlinkData();
    }

    public ExcelStyle createExcelStyle() {
        return new ExcelStyle();
    }

    public SheetStyle createSheetStyle() {
        return new SheetStyle();
    }

    public RowStyle createRowStyle() {
        return new RowStyle();
    }
}