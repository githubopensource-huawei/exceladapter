/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style;

import com.app.excelcsvconverter.xmlmodel.styledata.CellStyleType;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 基础样式生成器
 *
 * @since 2019-08-08
 */
public class CellStyleCreator {
    public static final short GROUP_TITLE_COLOR = 41;

    public static final short COLUMN_COLOR = 47;

    public static final String SUMMARY_FONT_NAME = "Arial";

    public static final short SUMMARY_TEXT_FONT_SIZE = 10;

    private CellStyle headStyle;

    private CellStyle columnStyle;

    private CellStyle dataStyle;

    private CellStyle hyperLinkStyle;

    private CellStyle defaultColumnStyle;

    private CellStyle defaultIubColumnStyle;

    private CellStyle iubGroupStyle;

    private CellStyle coverHeadStyle;

    private CellStyle coverVersionStyle;

    private CellStyle titleHyperLink;

    private CellStyle conflictHyperLinkStyle;

    private CellStyle featureHyperLink;

    private CellStyle summaryDataCenterLeft;

    private CellStyle summaryDataTopLeft;

    private DataFormat dataFormat;

    private Font commentFont;

    private Workbook workBook;

    public CellStyleCreator(Workbook workBook) {
        this.dataFormat = workBook.createDataFormat();
        this.workBook = workBook;
    }

    public DataFormat getDataFormat() {
        return this.dataFormat;
    }

    public CellStyle getStyleByType(CellStyleType type) {
        switch (type) {
            case HEAD_STYLE:
                return this.getSummaryHeadStyle();
            case COLUMN_STYLE:
                return this.getSummaryColumnStyle();
            case DATA_STYLE:
                return this.getSummaryDataStyle();
            case HYPER_LINK_STYLE:
                return this.getSummaryHyperLinkStyle();
            case IUB_GROUP_STYLE:
                return this.getIubGroupStyle();
            case IUB_COLUMN_STYLE:
                return this.getDefaultIubColumnStyle();
            case SUMMARY_PATTEN_HYPERLINK:
                return this.getSummaryHyperLinkStyle();
            case SUMMARY_TITLE_HYPERLINK:
                return this.getSummaryTitleHyperLinkStyle();
            case CONFLICT_LINK:
                return this.createConflictHyperLinkStyle();
            case FEATURE_HYPERLINK:
                return this.createFeatureHyperLink();
            case SUMMARY_DATA_TOP_LEFT:
                return this.createSummaryDataTopLeftStyle();
            case SUMMARY_DATA_CENTER_LEFT:
                return this.createSummaryDataCenterLeftStyle();
            default:
                return this.getDefaultStyle();
        }
    }

    public Font getFontByType(CellStyleType type) {
        switch (type) {
            case COMMENT_STYLE:
                return this.getCommentFont();
            default:
                return this.getEmptyFont();
        }
    }

    public CellStyle getEmptyCellStyle() {
        return this.workBook.createCellStyle();
    }

    public Font getEmptyFont() {
        return this.workBook.createFont();
    }

    public CellStyle getSummaryHeadStyle() {
        if (headStyle == null) {
            headStyle = this.createHeadStyle();
        }
        return headStyle;
    }

    public CellStyle getSummaryColumnStyle() {
        if (columnStyle == null) {
            columnStyle = this.createColumnStyle();
        }
        return columnStyle;
    }

    public CellStyle getSummaryDataStyle() {
        if (dataStyle == null) {
            dataStyle = createStyleWithBorderAlignTextAndSummaryFont(HorizontalAlignment.CENTER);
        }
        return dataStyle;
    }

    public CellStyle getSummaryHyperLinkStyle() {
        if (hyperLinkStyle == null) {
            hyperLinkStyle = this.createHyperLinkStyle();
        }
        return hyperLinkStyle;
    }

    public CellStyle getDefaultStyle() {
        if (defaultColumnStyle == null) {
            defaultColumnStyle = this.createDefaultStyle();
        }
        return defaultColumnStyle;
    }

    private CellStyle createHeadStyle() {
        CellStyle style = getCellStyleWithBorderSolidAlignCenterAndText(GROUP_TITLE_COLOR);
        Font font = getFont(SUMMARY_TEXT_FONT_SIZE, true);
        style.setFont(font);
        return style;
    }

    private CellStyle createColumnStyle() {
        CellStyle style = getCellStyleWithBorderSolidAlignCenterAndText(COLUMN_COLOR);
        Font font = getFont(SUMMARY_TEXT_FONT_SIZE, false);
        style.setFont(font);
        return style;
    }

    private CellStyle getCellStyleWithBorderSolidAlignCenterAndText(short colorIndex) {
        CellStyle style = createCellStyleWithBorder();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFillForegroundColor(colorIndex);
        style.setDataFormat(dataFormat.getFormat("@"));
        return style;
    }

    private CellStyle createCellStyleWithBorder() {
        CellStyle style = workBook.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        return style;
    }

    private Font getFont(short fontName, boolean isBold) {
        Font font = workBook.createFont();
        font.setFontName(SUMMARY_FONT_NAME);
        font.setFontHeightInPoints(fontName);
        font.setBold(isBold);
        return font;
    }

    private CellStyle createStyleWithBorderAlignTextAndSummaryFont(HorizontalAlignment align) {
        CellStyle style = createCellStyleWithBorder();
        style.setAlignment(align);
        style.setDataFormat(dataFormat.getFormat("@"));
        Font font = workBook.createFont();
        font.setFontName(SUMMARY_FONT_NAME);
        font.setFontHeightInPoints(SUMMARY_TEXT_FONT_SIZE);
        style.setFont(font);
        return style;
    }

    private CellStyle createDefaultStyle() {
        CellStyle style = workBook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(dataFormat.getFormat("@"));
        Font font = workBook.createFont();
        font.setFontName(SUMMARY_FONT_NAME);
        font.setFontHeightInPoints(SUMMARY_TEXT_FONT_SIZE);
        style.setFont(font);
        return style;
    }

    private CellStyle createHyperLinkStyle() {
        CellStyle hlinkStyle = createCellStyleWithBorder();
        hlinkStyle.setAlignment(HorizontalAlignment.LEFT);
        hlinkStyle.setFont(createLinkFont());
        return hlinkStyle;
    }

    private Font createLinkFont() {
        Font hlinkFont = workBook.createFont();
        hlinkFont.setUnderline(Font.U_SINGLE);
        hlinkFont.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
        return hlinkFont;
    }

    public CellStyle getSummaryCoverHeadStyle() {
        if (this.coverHeadStyle == null) {
            this.coverHeadStyle = this.createCoverTitleStyle();
        }

        return this.coverHeadStyle;
    }

    public void setSummaryCoverHeadStyle(CellStyle style) {
        this.coverHeadStyle = style;
    }

    private CellStyle createCoverTitleStyle() {
        CellStyle style = this.createCellStyleWithBorder();
        style.setFillForegroundColor((short) 47);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(this.dataFormat.getFormat("@"));
        Font font = this.workBook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        return style;
    }

    public CellStyle getSummaryCoverVersionStyle() {
        if (this.coverVersionStyle == null) {
            this.coverVersionStyle = this.createCoverVersionStyle();
        }

        return this.coverVersionStyle;
    }

    public void setSummaryCoverVersionStyle(CellStyle style) {
        this.coverVersionStyle = style;
    }

    private CellStyle createCoverVersionStyle() {
        CellStyle style = this.createCellStyleWithBorder();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setDataFormat(this.dataFormat.getFormat("@"));
        Font font = this.workBook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        return style;
    }

    public Font getCommentFont() {
        if (this.commentFont == null) {
            this.commentFont = this.createCommentFont();
        }

        return this.commentFont;
    }

    private Font createCommentFont() {
        Font font = this.workBook.createFont();
        font.setFontName("Tahoma");
        font.setBold(true);
        font.setFontHeightInPoints((short) 9);
        return font;
    }

    public CellStyle getDefaultIubColumnStyle() {
        if (this.defaultIubColumnStyle == null) {
            this.defaultIubColumnStyle = this.createDefaultIubColumnStyle();
        }

        return this.createDefaultIubColumnStyle();
    }

    private CellStyle createDefaultIubColumnStyle() {
        CellStyle style = this.workBook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(this.dataFormat.getFormat("@"));
        style.setFillForegroundColor((short) 22);
        style.setFillBackgroundColor((short) 22);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        Font font = this.workBook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        return style;
    }

    public CellStyle getIubGroupStyle() {
        if (this.iubGroupStyle == null) {
            this.iubGroupStyle = this.createIubGroupStyle();
        }

        return this.iubGroupStyle;
    }

    private CellStyle createIubGroupStyle() {
        CellStyle style = this.createCellStyleWithBorder();
        style.setFillBackgroundColor((short) 16);
        style.setFillPattern(FillPatternType.LEAST_DOTS);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setDataFormat(this.dataFormat.getFormat("@"));
        Font font = this.workBook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);
        font.setBold(true);
        font.setColor((short) 9);
        style.setFont(font);
        return style;
    }

    private CellStyle getSummaryTitleHyperLinkStyle() {
        if (this.titleHyperLink == null) {
            this.titleHyperLink = this.createColumnStyle();
            this.titleHyperLink.setFont(this.createLinkFont());
        }

        return this.titleHyperLink;
    }

    private CellStyle createConflictHyperLinkStyle() {
        if (this.conflictHyperLinkStyle == null) {
            this.conflictHyperLinkStyle = this.createHyperLinkStyle();
            this.conflictHyperLinkStyle.setAlignment(HorizontalAlignment.LEFT);
            this.conflictHyperLinkStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        }

        return this.conflictHyperLinkStyle;
    }

    private CellStyle createFeatureHyperLink() {
        if (this.featureHyperLink == null) {
            this.featureHyperLink = this.createHyperLinkStyle();
            this.featureHyperLink.setAlignment(HorizontalAlignment.CENTER);
        }

        return this.featureHyperLink;
    }

    private CellStyle createSummaryDataTopLeftStyle() {
        if (this.summaryDataTopLeft == null) {
            this.summaryDataTopLeft = this.createCellStyleWithBorder();
            this.summaryDataTopLeft.setAlignment(HorizontalAlignment.LEFT);
            this.summaryDataTopLeft.setVerticalAlignment(VerticalAlignment.TOP);
        }

        return this.summaryDataTopLeft;
    }

    private CellStyle createSummaryDataCenterLeftStyle() {
        if (this.summaryDataCenterLeft == null) {
            this.summaryDataCenterLeft = this.createCellStyleWithBorder();
            this.summaryDataCenterLeft.setAlignment(HorizontalAlignment.LEFT);
            this.summaryDataCenterLeft.setVerticalAlignment(VerticalAlignment.CENTER);
        }

        return this.summaryDataCenterLeft;
    }
}