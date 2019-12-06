/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.genexcelstyle;

import com.app.excelcsvconverter.xmlmodel.styledata.CellLevelStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.CellRange;
import com.app.excelcsvconverter.xmlmodel.styledata.CellStyleData;
import com.app.excelcsvconverter.xmlmodel.styledata.CellStyleType;
import com.app.excelcsvconverter.xmlmodel.styledata.ColumnStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.CommentArea;
import com.app.excelcsvconverter.xmlmodel.styledata.CommentData;
import com.app.excelcsvconverter.xmlmodel.styledata.ExcelStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.FontData;
import com.app.excelcsvconverter.xmlmodel.styledata.HyperlinkData;
import com.app.excelcsvconverter.xmlmodel.styledata.RowStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetVisible;
import com.app.excelcsvconverter.xmlmodel.styledata.StyleData;
import com.app.excelcsvconverter.xmlmodel.styledata.Validation;
import com.google.common.base.Strings;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STUnsignedIntHex;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 读取excel样式
 *
 * @since 2019-09-11
 */
final class ExcelReaderWithPOI {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelReaderWithPOI.class);

    private static final Set<String> SPECIAL_SHEETNAME_SET = new HashSet<String>() {{
        add("Qos");
    }};

    private static Pattern STR_PATTERN = Pattern.compile("[a-zA-Z]+");

    private static Pattern NUM_PATTERN = Pattern.compile("[0-9]+");

    private static Pattern NOT_NUM_PATTERN = Pattern.compile("[^0-9]+");

    private static Pattern ADDRESS_PATTERN = Pattern.compile("[R|r][0-9]+[C|c][0-9]+");

    private Workbook workbook;

    private List<CellStyleData> cellStyleDataList;

    private List<FontData> fontDataList;

    ExcelStyle getExcelStyle(File excelFile) {
        ExcelStyle excelStyle = new ExcelStyle();
        try (InputStream inputStream = new FileInputStream(excelFile)) {
            workbook = WorkbookFactory.create(inputStream);
            cellStyleDataList = excelStyle.getCellStyleData();
            fontDataList = excelStyle.getFontData();
            XSSFFont font = (XSSFFont) workbook.getFontAt(0);
            if (font != null) {
                excelStyle.setDefaultFontName(font.getFontName());
                excelStyle.setDefaultFontSize(font.getFontHeightInPoints());
            }
            excelStyle.setActiveSheet(workbook.getSheetName(workbook.getActiveSheetIndex()));
            excelStyle.setSortedSheetNames(getSortedSheetNames());
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                excelStyle.getSheetStyle().add(getSheetStyle(sheet, i));
            }
        } catch (FileNotFoundException e) {
            LOGGER.error("file {} not found. ", excelFile.getName());
        } catch (IOException e) {
            LOGGER.error("get excel style error.", e);
        }
        return excelStyle;
    }

    private SheetStyle getSheetStyle(Sheet sheet, int i) {
        SheetStyle sheetStyle = new SheetStyle();
        // set sheetName
        sheetStyle.setSheetName(sheet.getSheetName());
        // set displayGridLines
        sheetStyle.setDisplayGridLines(sheet.isDisplayGridlines());
        // set visible
        if (workbook.isSheetHidden(i)) {
            sheetStyle.setVisible(SheetVisible.HIDDEN);
            if (!SPECIAL_SHEETNAME_SET.contains(workbook.getSheetName(i))) {
                return sheetStyle;
            }
        } else if (workbook.isSheetVeryHidden(i)) {
            sheetStyle.setVisible(SheetVisible.VERY_HIDDEN);
            return sheetStyle;
        } else {
            sheetStyle.setVisible(SheetVisible.VISIBLE);
        }
        // set Validation
        setValidationData(sheet, sheetStyle);
        // set CellRange
        setCellRangeData(sheet, sheetStyle);
        // set rowStyle
        setRowStyle(sheetStyle, sheet);
        // set ColumnStyle
        setColumnStyle(sheetStyle, sheet);
        // set CellLevelStyle
        sheetStyle.setCellLevelStyle(getCellStyle(sheet));
        // merge cell
        mergeCellStyle(sheetStyle.getCellLevelStyle().getStyleData());
        // set tabColor
        setTabColorData(sheet, sheetStyle);
        return sheetStyle;
    }

    private void setValidationData(Sheet sheet, SheetStyle sheetStyle) {
        for (DataValidation validation : sheet.getDataValidations()) {
            CellRangeAddress[] cellRangeAddresses = validation.getRegions().getCellRangeAddresses();
            for (int i = 0; i < cellRangeAddresses.length; i++) {
                Validation validation1 = new Validation();
                CellRange cellRange = new CellRange();
                cellRange.setRowEnd(cellRangeAddresses[i].getLastRow());
                cellRange.setRowBegin(cellRangeAddresses[i].getFirstRow());
                cellRange.setColEnd(cellRangeAddresses[i].getLastColumn());
                cellRange.setColBegin(cellRangeAddresses[i].getFirstColumn());
                validation1.setPromptBoxTitle(Strings.nullToEmpty(validation.getPromptBoxTitle()));
                validation1.setPromptBoxText(Strings.nullToEmpty(validation.getPromptBoxText()));
                validation1.setErrorBoxTitle(Strings.nullToEmpty(validation.getErrorBoxTitle()));
                validation1.setErrorBoxText(Strings.nullToEmpty(validation.getErrorBoxText()));
                validation1.setCellRange(cellRange);
                validation1.setAttrs(getAttrs(validation));
                validation1.setAttrs2(getAttrs2(validation));
                validation1.setValidationType(validation.getValidationConstraint().getValidationType());
                validation1.setOperator(validation.getValidationConstraint().getOperator());
                sheetStyle.getValidation().add(validation1);
            }
        }
    }

    private void setCellRangeData(Sheet sheet, SheetStyle sheetStyle) {
        for (CellRangeAddress cellAddresses : sheet.getMergedRegions()) {
            CellRange cellRange = new CellRange();
            cellRange.setRowEnd(cellAddresses.getLastRow());
            cellRange.setRowBegin(cellAddresses.getFirstRow());
            cellRange.setColEnd(cellAddresses.getLastColumn());
            cellRange.setColBegin(cellAddresses.getFirstColumn());
            sheetStyle.getCellRange().add(cellRange);
        }
    }

    private void setTabColorData(Sheet sheet, SheetStyle sheetStyle) {
        XSSFColor tabColor = ((XSSFSheet) sheet).getTabColor();
        if (tabColor != null && tabColor.getRGB() != null) {
            String tableColor = getRgbColor(tabColor.getRGB());
            sheetStyle.setTabColor(tableColor);
        }
    }

    private String getRgbColor(byte[] rgb) {
        String rgbColor = "";
        if (null != rgb) {
            int red = rgb[0] >= 0 ? rgb[0] : 255 + rgb[0];
            int green = rgb[1] >= 0 ? rgb[1] : 255 + rgb[1];
            int black = rgb[2] >= 0 ? rgb[2] : 255 + rgb[2];
            rgbColor = red + "," + green + "," + black;
        }
        return rgbColor;
    }

    private void mergeCellStyle(List<StyleData> styleDataList) {
        if (styleDataList.size() < 2) {
            return;
        }
        List<StyleData> styleDataMergedList = new ArrayList<>();
        for (int i = 0; i < styleDataList.size(); i++) {
            StyleData lastStyleData = styleDataList.get(i);
            for (int j = i + 1; j < styleDataList.size(); j++) {
                StyleData styleData = styleDataList.get(j);
                if (isSameStyleData(lastStyleData, styleData) && isSameColRange(lastStyleData, styleData)) {
                    CellRange lastCellRange = lastStyleData.getCellRange();
                    lastCellRange.setRowEnd(styleData.getCellRange().getRowEnd());
                    styleDataMergedList.add(styleData);
                }
            }
        }
        styleDataList.removeAll(styleDataMergedList);
    }

    private boolean isSameStyleData(StyleData styleData1, StyleData styleData2) {
        if (styleData1 == null || styleData2 == null) {
            return false;
        }
        if (styleData1.getCellStyleIndex() != styleData2.getCellStyleIndex()) {
            return false;
        }
        return styleData1.getFontIndex() == styleData2.getFontIndex();
    }

    private boolean isSameColRange(StyleData lastStyleData, StyleData styleData) {
        CellRange lastCellRange = lastStyleData.getCellRange();
        CellRange cellRange = styleData.getCellRange();
        if (lastCellRange.getColBegin() != cellRange.getColBegin()) {
            return false;
        }
        if (lastCellRange.getColEnd() != cellRange.getColEnd()) {
            return false;
        }
        return lastCellRange.getRowEnd() + 1 == cellRange.getRowBegin();
    }

    private String getAttrs(DataValidation validation) {
        if (!Strings.isNullOrEmpty(validation.getValidationConstraint().getFormula1())) {
            int validationType = validation.getValidationConstraint().getValidationType();
            String formula1 = validation.getValidationConstraint().getFormula1();
            if (validationType == DataValidationConstraint.ValidationType.FORMULA || (
                validationType == DataValidationConstraint.ValidationType.LIST && formula1.contains(":"))) {
                return formula1;
            }
            return formula1.replaceAll("\"", "");
        }
        return "";
    }

    private String getAttrs2(DataValidation validation) {
        if (!Strings.isNullOrEmpty(validation.getValidationConstraint().getFormula2())) {
            return validation.getValidationConstraint().getFormula2().replaceAll("\"", "");
        }
        return "";
    }

    private void setRowStyle(SheetStyle sheetStyle, Sheet sheet) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (null == row) {
                continue;
            }
            RowStyle rowStyle = new RowStyle();
            rowStyle.setHeight((int) row.getHeight());
            rowStyle.setRowBegin(row.getRowNum());
            rowStyle.setRowEnd(row.getRowNum());
            if (null != ((XSSFRow) row).getCTRow()) {
                boolean isHidden = ((XSSFRow) row).getCTRow().getHidden();
                rowStyle.setHidden(isHidden);
            }
            List<RowStyle> rowStyleList = sheetStyle.getRowStyle();
            if (rowStyleList.isEmpty()) {
                rowStyleList.add(rowStyle);
            } else {
                RowStyle lastRowStyle = rowStyleList.get(rowStyleList.size() - 1);
                boolean isSame = isSameRowStyle(lastRowStyle, rowStyle);
                if (isSame) {
                    lastRowStyle.setRowEnd(rowStyle.getRowEnd());
                } else {
                    rowStyleList.add(rowStyle);
                }
            }
        }
    }

    private boolean isSameRowStyle(RowStyle rowStyle1, RowStyle rowStyle2) {
        int height1 = rowStyle1.getHeight() == null ? 0 : rowStyle1.getHeight();
        int height2 = rowStyle2.getHeight() == null ? 0 : rowStyle2.getHeight();
        if (height1 != height2) {
            return false;
        }
        boolean isHidden1 = rowStyle1.isHidden() == null ? false : rowStyle1.isHidden();
        boolean isHidden2 = rowStyle2.isHidden() == null ? false : rowStyle2.isHidden();
        return isHidden1 == isHidden2;
    }

    private void setColumnStyle(SheetStyle sheetStyle, Sheet sheet) {
        int maxColumnNum = 1;
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (null == row) {
                continue;
            }
            maxColumnNum = Math.max(maxColumnNum, row.getLastCellNum());
        }

        for (int i = 0; i < maxColumnNum; i++) {
            ColumnStyle columnStyle = new ColumnStyle();
            columnStyle.setColBegin(i);
            columnStyle.setColEnd(i);
            columnStyle.setHidden(sheet.isColumnHidden(i));
            columnStyle.setWidth(sheet.getColumnWidth(i));
            List<ColumnStyle> columnStyleList = sheetStyle.getColumnStyle();
            if (columnStyleList.isEmpty()) {
                columnStyleList.add(columnStyle);
            } else {
                ColumnStyle lastColumnStyle = columnStyleList.get(columnStyleList.size() - 1);
                boolean isSame = isSameColumnStyle(lastColumnStyle, columnStyle);
                if (isSame) {
                    lastColumnStyle.setColEnd(columnStyle.getColEnd());
                } else {
                    columnStyleList.add(columnStyle);
                }
            }
        }
    }

    private boolean isSameColumnStyle(ColumnStyle columnStyle1, ColumnStyle columnStyle2) {
        int width1 = columnStyle1.getWidth() == null ? 0 : columnStyle1.getWidth();
        int width2 = columnStyle2.getWidth() == null ? 0 : columnStyle2.getWidth();
        if (width1 != width2) {
            return false;
        }
        boolean isHidden1 = columnStyle1.isHidden() == null ? false : columnStyle1.isHidden();
        boolean isHidden2 = columnStyle2.isHidden() == null ? false : columnStyle2.isHidden();
        return isHidden1 == isHidden2;
    }

    private CellLevelStyle getCellStyle(Sheet sheet) {
        CellLevelStyle cellLevelStyle = new CellLevelStyle();
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (null == row) {
                continue;
            }
            int index = row.getLastCellNum() - 1;
            for (int j = 0; j <= index; j++) {
                Cell cell = row.getCell(j);
                if (null == cell) {
                    continue;
                }
                this.initCellLevelStyle(sheet.getSheetName(), cell.getRowIndex() + 1,
                    CellReference.convertNumToColString(cell.getColumnIndex()), cellLevelStyle);
            }
        }
        return cellLevelStyle;
    }

    private String getSortedSheetNames() {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            sb.append(workbook.getSheetAt(i).getSheetName()).append(",");
        }

        if (!sb.toString().isEmpty()) {
            sb.delete(sb.length() - 1, sb.length());
        }
        return sb.toString();
    }

    /**
     * 获取单元格值
     *
     * @param sheetName sheet页名称, 不支持null或空字符
     * @param rowId     单元格行号，从1开始
     * @param columnId  单元格列号，从'A'开始
     */
    private void initCellLevelStyle(String sheetName, int rowId, String columnId, CellLevelStyle cellLevelStyle) {
        if (Strings.isNullOrEmpty(sheetName)) {
            return;
        }
        if (rowId < 1) {
            return;
        }
        // 单元格行号，从1开始, 但POI从0开始，所以要--
        rowId--;
        int columnNum = CellReference.convertColStringToIndex(columnId);
        if (columnNum < 0) {
            return;
        }
        Sheet sheet = workbook.getSheet(sheetName);
        if (null == sheet) {
            return;
        }
        Row row = sheet.getRow(rowId);
        if (null == row) {
            return;
        }
        Cell cell = row.getCell(columnNum);
        if (null == cell) {
            return;
        }
        initALineCellStyleToList(row, columnNum, columnNum, cellLevelStyle);
    }

    private void initALineCellStyleToList(Row row, int columnStartIndex, int columnEndIndex,
        CellLevelStyle cellLevelStyle) {
        int index = columnEndIndex < 0 ? row.getLastCellNum() - 1 : columnEndIndex;

        for (int j = columnStartIndex; j <= index; j++) {
            Cell cell = row.getCell(j);
            if (null == cell) {
                continue;
            }
            // 添加批注样式
            Comment comment = cell.getCellComment();
            if (null != comment) {
                cellLevelStyle.getCommentData().add(getCommentData(cell, comment));
            }
            // 添加超链接样式
            Hyperlink hyperlink = cell.getHyperlink();
            if (null != hyperlink) {
                cellLevelStyle.getHyperlinkData().add(getHyperLinkData(cell, hyperlink));
            }
            // 添加单元格样式
            List<StyleData> styleDataList = cellLevelStyle.getStyleData();
            // 合并行单元格样式
            mergeALineCellStyle(styleDataList, getStyleData(cell));
        }
    }

    private void mergeALineCellStyle(List<StyleData> styleDataList, StyleData styleData) {
        if (styleDataList.size() < 2) {
            styleDataList.add(styleData);
            return;
        }
        StyleData lastStyleData = styleDataList.get(styleDataList.size() - 1);

        if (isSameStyleData(lastStyleData, styleData) && isSameLine(lastStyleData, styleData)) {
            CellRange lastCellRange = lastStyleData.getCellRange();
            lastCellRange.setColEnd(styleData.getCellRange().getColEnd());
        } else {
            styleDataList.add(styleData);
        }
    }

    private boolean isSameLine(StyleData lastStyleData, StyleData styleData) {
        CellRange lastCellRange = lastStyleData.getCellRange();
        CellRange cellRange = styleData.getCellRange();
        if (lastCellRange.getRowBegin() != cellRange.getRowBegin()) {
            return false;
        }
        if (lastCellRange.getRowEnd() != cellRange.getRowEnd()) {
            return false;
        }
        return lastCellRange.getColEnd() + 1 == cellRange.getColBegin();
    }

    private boolean isSameCellStyleData(CellStyleData cellStyleData1, CellStyleData cellStyleData2) {
        if (cellStyleData1 == null || cellStyleData2 == null) {
            return false;
        }
        if (!Strings.nullToEmpty(cellStyleData1.getBackgroundColor())
            .equals(Strings.nullToEmpty(cellStyleData2.getBackgroundColor()))) {
            return false;
        }
        short dataFormat1 = cellStyleData1.getDataFormat() == null ? -1 : cellStyleData1.getDataFormat();
        short dataFormat2 = cellStyleData2.getDataFormat() == null ? -1 : cellStyleData2.getDataFormat();
        if (dataFormat1 != dataFormat2) {
            return false;
        }
        if (!Strings.nullToEmpty(cellStyleData1.getDataFormatString())
            .equals(Strings.nullToEmpty(cellStyleData2.getDataFormatString()))) {
            return false;
        }
        if (!Strings.nullToEmpty(cellStyleData1.getAlignment())
            .equals(Strings.nullToEmpty(cellStyleData2.getAlignment()))) {
            return false;
        }
        if (!Strings.nullToEmpty(cellStyleData1.getVerticalAlignment())
            .equals(Strings.nullToEmpty(cellStyleData2.getVerticalAlignment()))) {
            return false;
        }
        return isSameCellStyle(cellStyleData1, cellStyleData2);
    }

    private boolean isSameCellStyle(CellStyleData cellStyleData1, CellStyleData cellStyleData2) {
        if (!Strings.nullToEmpty(cellStyleData1.getBorderBottom())
            .equals(Strings.nullToEmpty(cellStyleData2.getBorderBottom()))) {
            return false;
        }
        if (!Strings.nullToEmpty(cellStyleData1.getBorderLeft())
            .equals(Strings.nullToEmpty(cellStyleData2.getBorderLeft()))) {
            return false;
        }
        if (!Strings.nullToEmpty(cellStyleData1.getBorderTop())
            .equals(Strings.nullToEmpty(cellStyleData2.getBorderTop()))) {
            return false;
        }
        if (!Strings.nullToEmpty(cellStyleData1.getBorderRight())
            .equals(Strings.nullToEmpty(cellStyleData2.getBorderRight()))) {
            return false;
        }
        if (!Strings.nullToEmpty(cellStyleData1.getBorderRight())
            .equals(Strings.nullToEmpty(cellStyleData2.getBorderRight()))) {
            return false;
        }
        boolean wrapText1 = cellStyleData1.isWrapText() == null ? false : cellStyleData1.isWrapText();
        boolean wrapText2 = cellStyleData2.isWrapText() == null ? false : cellStyleData2.isWrapText();
        if (wrapText1 != wrapText2) {
            return false;
        }
        boolean isLocked1 = cellStyleData1.isLocked() == null ? false : cellStyleData1.isLocked();
        boolean isLocked2 = cellStyleData2.isLocked() == null ? false : cellStyleData2.isLocked();
        return isLocked1 == isLocked2;
    }

    private boolean isSameFontData(FontData fontData1, FontData fontData2) {
        if (fontData1 == null || fontData2 == null) {
            return false;
        }
        if (!Strings.nullToEmpty(fontData1.getFontName()).equals(Strings.nullToEmpty(fontData2.getFontName()))) {
            return false;
        }
        short fontSize1 = fontData1.getFontSize() == null ? -1 : fontData1.getFontSize();
        short fontSize2 = fontData2.getFontSize() == null ? -1 : fontData2.getFontSize();
        if (fontSize1 != fontSize2) {
            return false;
        }
        if (!Strings.nullToEmpty(fontData1.getColor()).equals(Strings.nullToEmpty(fontData2.getColor()))) {
            return false;
        }
        boolean bold1 = fontData1.isBold() == null ? false : fontData1.isBold();
        boolean bold2 = fontData2.isBold() == null ? false : fontData2.isBold();
        if (bold1 != bold2) {
            return false;
        }
        byte underline1 = fontData1.getUnderline() == null ? -1 : fontData1.getUnderline();
        byte underline2 = fontData2.getUnderline() == null ? -1 : fontData2.getUnderline();
        return underline1 == underline2;
    }

    private StyleData getStyleData(Cell cell) {
        StyleData styleData = new StyleData();
        CellStyleData cellStyleData = getCellStyleData(cell);
        styleData.setCellStyleIndex(getCellStyleIndex(cellStyleData));
        FontData fontData = getFontData(cell);
        styleData.setFontIndex(getFontIndex(fontData));
        styleData.setCellStyleType(CellStyleType.CUSTOMIZED);

        CellRange cellRange = new CellRange();
        cellRange.setColBegin(cell.getColumnIndex());
        cellRange.setColEnd(cell.getColumnIndex());
        cellRange.setRowBegin(cell.getRowIndex());
        cellRange.setRowEnd(cell.getRowIndex());
        styleData.setCellRange(cellRange);
        return styleData;
    }

    private HyperlinkData getHyperLinkData(Cell cell, Hyperlink hyperlink) {
        HyperlinkData hyperlinkData = new HyperlinkData();
        hyperlinkData.setCellStyleType(CellStyleType.CUSTOMIZED);
        hyperlinkData.setColumn(cell.getColumnIndex());
        hyperlinkData.setRow(cell.getRowIndex());
        FontData fontData = getFontData(cell);
        hyperlinkData.setFontIndex(getFontIndex(fontData));
        CellStyleData cellStyleData = getCellStyleData(cell);
        hyperlinkData.setCellStyleIndex(getCellStyleIndex(cellStyleData));

        String address = hyperlink.getAddress();
        if (!Strings.isNullOrEmpty(address) && address.contains("!")) {
            String[] datas = address.split("!");
            hyperlinkData.setLinkSheetName(datas[0].replaceAll("'", ""));
            Matcher hyperlinkType = ADDRESS_PATTERN.matcher(datas[1]);
            if (hyperlinkType.find()) {
                // 链接格式：R1C1
                Matcher matcher = NOT_NUM_PATTERN.matcher(datas[1]);
                String[] result = matcher.replaceAll(" ").trim().split(" ");
                hyperlinkData.setLinkRow(Integer.parseInt(result[0]));
                hyperlinkData.setLinkColumn(Integer.parseInt(result[1]));
            } else {
                // 链接格式：A1
                String[] result = splitStrAndNumber(datas[1]);
                if (result[0] == null) {
                    return null;
                }
                hyperlinkData.setLinkColumn(CellReference.convertColStringToIndex(result[0]) + 1);
                int rowId = Integer.parseInt(result[1]);
                hyperlinkData.setLinkRow(rowId);
            }
        }
        return hyperlinkData;
    }

    private String[] splitStrAndNumber(String str) {
        String[] results = new String[2];
        Matcher matcher = STR_PATTERN.matcher(Strings.nullToEmpty(str));
        if (matcher.find()) {
            results[0] = matcher.group(0);
        }
        matcher = NUM_PATTERN.matcher(Strings.nullToEmpty(str));
        if (matcher.find()) {
            results[1] = matcher.group(0);
        }
        return results;
    }

    private CellStyleData getCellStyleData(Cell cell) {
        CellStyleData cellStyleData = new CellStyleData();
        cellStyleData.setBackgroundColor(getColor(cell));
        cellStyleData.setWrapText(cell.getCellStyle().getWrapText());
        cellStyleData.setDataFormat(cell.getCellStyle().getDataFormat());
        cellStyleData.setDataFormatString(cell.getCellStyle().getDataFormatString());
        cellStyleData.setAlignment(cell.getCellStyle().getAlignment().toString());
        cellStyleData.setBorderBottom(cell.getCellStyle().getBorderBottom().toString());
        cellStyleData.setBorderLeft(cell.getCellStyle().getBorderLeft().toString());
        cellStyleData.setBorderRight(cell.getCellStyle().getBorderRight().toString());
        cellStyleData.setBorderTop(cell.getCellStyle().getBorderTop().toString());
        cellStyleData.setVerticalAlignment(cell.getCellStyle().getVerticalAlignment().toString());
        cellStyleData.setFillPattern(cell.getCellStyle().getFillPattern().toString());
        cellStyleData.setLocked(cell.getCellStyle().getLocked());
        return cellStyleData;
    }

    private int getCellStyleIndex(CellStyleData cellStyleData) {
        for (CellStyleData cellStyle : cellStyleDataList) {
            boolean result = isSameCellStyleData(cellStyleData, cellStyle);
            if (result) {
                return cellStyle.getIndex();
            }
        }
        cellStyleDataList.add(cellStyleData);
        cellStyleData.setIndex(cellStyleDataList.indexOf(cellStyleData));
        return cellStyleData.getIndex();
    }

    private String getColor(Cell cell) {
        XSSFColor foregroundColor = (XSSFColor) cell.getCellStyle().getFillForegroundColorColor();
        XSSFColor backgroundColor = (XSSFColor) cell.getCellStyle().getFillBackgroundColorColor();
        if (null != foregroundColor) {
            String color = getColorString(foregroundColor);
            if (color != null) {
                return color;
            }
        } else if (backgroundColor != null) {
            String color = getColorString(backgroundColor);
            if (color != null) {
                return color;
            }
        }
        return "";
    }

    private String getColorString(XSSFColor xssfColor) {
        byte[] b = xssfColor.getRGBWithTint();

        if (b != null) {
            return getRgbColor(b);
        }

        b = xssfColor.getRGB();
        if (b != null) {
            return getRgbColor(b);
        }
        return null;
    }

    private CommentData getCommentData(Cell cell, Comment comment) {
        CommentData commentData = new CommentData();
        // set font data
        FontData fontData = getFontData(comment);
        commentData.setFontIndex(getFontIndex(fontData));
        // 从0开始数
        commentData.setRow(cell.getRowIndex());
        // 从0开始数
        commentData.setColumn(cell.getColumnIndex());
        //set comment
        if (Strings.isNullOrEmpty(comment.getString().getString())) {
            commentData.setComment("");
        } else {
            commentData.setComment(comment.getString().getString());
        }
        // set Comment Area
        CommentArea commentArea = new CommentArea();
        try {
            ClientAnchor clientAnchor = comment.getClientAnchor();
            if (clientAnchor != null) {
                commentArea.setDx1(clientAnchor.getDx1());
                commentArea.setDx2(clientAnchor.getDx2());
                commentArea.setDy1(clientAnchor.getDy1());
                commentArea.setDy2(clientAnchor.getDy2());
                commentArea.setCol1(cell.getColumnIndex());
                commentArea.setCol2((int) clientAnchor.getCol2());
                commentArea.setRow1(cell.getRowIndex());
                commentArea.setRow2(clientAnchor.getRow2());
            }
            commentData.setCommentArea(commentArea);
        } catch (Exception e) {
            LOGGER.error("Gets comment area value error", e);
        }
        commentData.setCellStyleType(CellStyleType.CUSTOMIZED);
        return commentData;
    }

    private FontData getFontData(Comment comment) {
        FontData fontData = new FontData();
        XSSFFont xssfFont = ((XSSFRichTextString) comment.getString()).getFontAtIndex(0);
        if (null == xssfFont) {
            return fontData;
        }
        fontData.setBold(xssfFont.getBold());
        fontData.setFontName(xssfFont.getFontName());
        fontData.setFontSize(xssfFont.getFontHeightInPoints());
        fontData.setUnderline(xssfFont.getUnderline());
        XSSFColor xssfColor = xssfFont.getXSSFColor();
        if (xssfColor == null) {
            return fontData;
        }
        CTColor ctColor = xssfColor.getCTColor();
        if (ctColor == null) {
            return fontData;
        }

        STUnsignedIntHex stUnsignedIntHex = ctColor.xgetRgb();
        if (null != stUnsignedIntHex) {
            String color = stUnsignedIntHex.getStringValue();
            int colorInt = (int) Long.parseLong(color, 16);
            Color c = new Color(colorInt);
            fontData.setColor(c.getRed() + "," + c.getGreen() + "," + c.getBlue());
        } else {
            byte[] b = xssfFont.getXSSFColor().getRGB();
            if (null != b) {
                fontData.setColor(getRgbColor(b));
            }
        }
        return fontData;
    }

    private FontData getFontData(Cell cell) {
        XSSFFont xssfFont = ((XSSFCell) cell).getCellStyle().getFont();
        FontData fontData = new FontData();
        fontData.setBold(xssfFont.getBold());

        fontData.setFontName(xssfFont.getFontName());
        fontData.setFontSize(xssfFont.getFontHeightInPoints());
        fontData.setUnderline(xssfFont.getUnderline());

        XSSFColor xssfColor = xssfFont.getXSSFColor();
        if (xssfColor == null) {
            return fontData;
        }
        CTColor ctColor = xssfColor.getCTColor();
        if (ctColor == null) {
            return fontData;
        }
        STUnsignedIntHex stUnsignedIntHex = ctColor.xgetRgb();
        if (null != stUnsignedIntHex) {
            String color = stUnsignedIntHex.getStringValue();
            int colorInt = (int) Long.parseLong(color, 16);
            Color c = new Color(colorInt);
            fontData.setColor(c.getRed() + "," + c.getGreen() + "," + c.getBlue());
        } else {
            byte[] b = xssfFont.getXSSFColor().getRGB();
            if (null != b) {
                fontData.setColor(getRgbColor(b));
            }
        }
        return fontData;
    }

    private int getFontIndex(FontData fontData) {
        for (FontData font : fontDataList) {
            boolean result = isSameFontData(fontData, font);
            if (result) {
                return font.getIndex();
            }
        }
        fontDataList.add(fontData);
        fontData.setIndex(fontDataList.indexOf(fontData));
        return fontData.getIndex();
    }
}