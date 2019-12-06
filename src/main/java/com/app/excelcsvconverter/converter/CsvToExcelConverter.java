/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.converter;

import static com.app.excelcsvconverter.businessoperation.summary.SummaryCoverSheet.COVER_START_ROW;
import static com.app.excelcsvconverter.businessoperation.summary.SummaryCoverSheet.MAX_TITLE_ROW;
import static com.app.excelcsvconverter.businessoperation.summary.SummaryCoverSheet.TITLE_ROW;
import static com.app.excelcsvconverter.businessoperation.summary.SummaryCoverSheet.TITLE_START_COL;
import static com.app.excelcsvconverter.businessoperation.summary.SummaryCoverSheet.TITLE_VALUE_COL;

import com.app.excelcsvconverter.businessoperation.summary.SummaryCoverSheet;
import com.app.excelcsvconverter.consts.summary.SummaryBookConst;
import com.app.excelcsvconverter.csv.CsvSheet;
import com.app.excelcsvconverter.parser.CloseableIterable;
import com.app.excelcsvconverter.read.JaxbReader;
import com.app.excelcsvconverter.resultmodel.ResultData;
import com.app.excelcsvconverter.style.CellStyleCreator;
import com.app.excelcsvconverter.style.cell.CellStyleFactory;
import com.app.excelcsvconverter.style.column.ColumnStyleFactory;
import com.app.excelcsvconverter.style.row.RowStyleFactory;
import com.app.excelcsvconverter.style.sheet.SheetStyleFactory;
import com.app.excelcsvconverter.style.workbook.WorkbookStyleFactory;
import com.app.excelcsvconverter.util.CompressUtil;
import com.app.excelcsvconverter.util.ConfigUtil;
import com.app.excelcsvconverter.util.ExcelConstants;
import com.app.excelcsvconverter.util.ExcelUtil;
import com.app.excelcsvconverter.util.FileUtil;
import com.app.excelcsvconverter.util.MessageUtil;
import com.app.excelcsvconverter.xmlmodel.styledata.CellLevelStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.CellRange;
import com.app.excelcsvconverter.xmlmodel.styledata.CellStyleData;
import com.app.excelcsvconverter.xmlmodel.styledata.CellStyleType;
import com.app.excelcsvconverter.xmlmodel.styledata.CommentData;
import com.app.excelcsvconverter.xmlmodel.styledata.ExcelStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.FontData;
import com.app.excelcsvconverter.xmlmodel.styledata.HyperlinkData;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.StyleData;
import com.google.common.base.Strings;
import com.google.common.collect.Iterables;
import com.google.common.collect.Sets;
import com.google.common.io.Files;

import java.awt.Color;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Csv To Excel转换器
 *
 * @since 2019-07-31
 */
public class CsvToExcelConverter {

    public static final String MACRO_TOOL_PATH = "Data" + File.separator;

    public static final String SUMMARYRES = "SummaryRes";

    private static final Logger LOGGER = LoggerFactory.getLogger(CsvToExcelConverter.class);

    private static final String CSV_SUFFIX_NAME = ".csv";

    private static final String SHEET_STYLE = "SHEET STYLE";

    private static final String STYLE_FILE_NAME = "style.xml";

    private static final String CSV_FILE = "CsvFile";

    /**
     * 存放模板key文件
     */
    private static final String FILEIDENTIFICATION_CSV_NAME = "FileIdentification";

    private final static int FLUSH_DATA_SIZE = 100;

    private final static int FONT_HEIGHT_MODULUS_A = 25;

    private final static int FONT_HEIGHT_MODULUS_B = 17;

    private final static int CHINESE_CHAR_MODULUS = 2;

    private final static int DEFAULT_COLUMN_WIDTH = 2048;

    private final static int COLUMN_WIDTH_MODULUS_A = 256;

    private final static int COLUMN_WIDTH_MODULUS_B = 184;

    private final String XLSX_SUFFIX_NAME = ".xlsx";

    private final String XLSM_SUFFIX_NAME = ".xlsm";

    private SXSSFWorkbook sxssfWorkbook;

    private ExcelStyle excelStyle = new ExcelStyle();

    private List<CsvSheet> csvSheetList = new ArrayList<>();

    private String language;

    private String temlpatePath;

    private List<String> fileIdList;

    /**
     * summary 标识
     */
    private String fileType;

    public CsvToExcelConverter(String language) {
        this.language = language;
    }

    /**
     * 解压Zip 初始化数据
     *
     * @param zipFile        zip文件
     * @param targetTempPath 缓存文件
     * @return ResultData初始化完成结果
     */
    private ResultData initData(File zipFile, String targetTempPath) {
        // 解压的csv文件路径
        String tempFilePath;
        try {
            tempFilePath = CompressUtil.unZipFiles(zipFile, targetTempPath);
        } catch (Exception e) {
            return new ResultData(false,
                MessageUtil.getMessage("COMPRESS_FILE_INCORRECT_EN", "COMPRESS_FILE_INCORRECT_CN", language));
        }
        File dir = new File(tempFilePath);
        File[] files = dir.listFiles();
        if (files == null || files.length == 0) {
            LOGGER.error("Gets csv file list is empty.");
            return new ResultData(false,
                MessageUtil.getMessage("UNZIP_DIRECTORY_EMPTY_EN", "UNZIP_DIRECTORY_EMPTY_CN", language));
        }
        for (File file : files) {
            if (file.isFile() && file.getName().endsWith(CSV_SUFFIX_NAME)) {
                String sheetName = file.getName().substring(0, file.getName().length() - 4);
                csvSheetList.add(new CsvSheet(sheetName, file.getParentFile()));
            } else if (file.isDirectory()) {
                File[] childFolder = file.listFiles();
                for (File subFile : childFolder) {
                    if (subFile.isFile() && subFile.getName().endsWith(CSV_SUFFIX_NAME)) {
                        String sheetName = subFile.getName().substring(0, subFile.getName().length() - 4);
                        if (!SHEET_STYLE.equals(sheetName)) {
                            csvSheetList.add(new CsvSheet(sheetName, subFile.getParentFile()));
                        }
                    } else if (subFile.isFile() && STYLE_FILE_NAME.equals(subFile.getName())) {
                        excelStyle = JaxbReader.readConfig(ExcelStyle.class, subFile);
                    }
                }
            }
        }
        return new ResultData(true, tempFilePath);
    }

    /**
     * 复制模板到缓存目录
     *
     * @param zipFile        zip文件
     * @param targetTempFile 缓存目录
     * @return
     */
    private ResultData copyTemplateFile(File zipFile, File targetTempFile) {
        String baseTemplatePath = temlpatePath;
        if (baseTemplatePath == null) {
            return new ResultData(false,
                MessageUtil.getMessage("BASE_TEMPLATE_PATH_EMPTY_EN", "BASE_TEMPLATE_PATH_EMPTY_CN", language));
        }
        File baseTemplateFile = new File(baseTemplatePath);
        if (!baseTemplateFile.exists() || !baseTemplateFile.isFile()) {
            return new ResultData(false,
                MessageUtil.getMessage("BASE_TEMPLATE_NOT_FOUND_EN", "BASE_TEMPLATE_NOT_FOUND_CN", language));
        }
        // Excel文件名
        String excelName = zipFile.getName().substring(0, zipFile.getName().indexOf("."));
        // Excel文件后缀名
        String suffixName = baseTemplatePath.substring(baseTemplatePath.indexOf("."));
        try {
            FileUtil.copyFileToDirectory(baseTemplateFile, targetTempFile, excelName + suffixName);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return new ResultData(true, targetTempFile.getPath() + File.separator + excelName + suffixName);
    }

    private ResultData initWorkbook(File excelFile) {
        try (BufferedInputStream bufIn = new BufferedInputStream(new FileInputStream(excelFile))) {
            OPCPackage pkg = OPCPackage.open(bufIn);
            XSSFWorkbook xssfWorkbook = new XSSFWorkbook(pkg);
            sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, FLUSH_DATA_SIZE);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            LOGGER.error("Init workbook fail.", e);
            return new ResultData(false,
                MessageUtil.getMessage("FAILED_LOAD_TEMPLATE_FILE_EN", "FAILED_LOAD_TEMPLATE_FILE_CN", language));
        } catch (InvalidFormatException e) {
            LOGGER.error("Init workbook fail.", e);
            return new ResultData(false,
                MessageUtil.getMessage("FAILED_LOAD_TEMPLATE_FILE_EN", "FAILED_LOAD_TEMPLATE_FILE_CN", language));
        }
        return new ResultData(true, "");
    }

    public ResultData converter(String zipFilePath) {
        File zipFile = new File(zipFilePath);
        if (!zipFile.exists() || !zipFile.isFile()) {
            return new ResultData(false, MessageUtil.getMessage("FILE_NOT_EXIST_EN", "FILE_NOT_EXIST_CN", language));
        }
        // 创建临时文件目录
        File tempDir = Files.createTempDir();
        String targetTempPath = tempDir + File.separator + CompressUtil.uniqueDirectoryName(CSV_FILE);
        File targetTempFile = new File(targetTempPath);
        if (!targetTempFile.mkdir()) {
            return new ResultData(false,
                MessageUtil.getMessage("GENERATE_TEMPORARY_FILE_EN", "GENERATE_TEMPORARY_FILE_CN", language));
        }
        // 解压zip文件并初始化数据对象
        ResultData initResult = initData(zipFile, targetTempPath);
        if (!initResult.isSuccessful()) {
            return initResult;
        }

        if (fileIdList != null && !fileIdList.isEmpty()) {
            temlpatePath = getTemplatePath(fileIdList);
        } else {
            temlpatePath = "";
        }
        boolean isTemlate = false;
        // 初始化workbook对象  临时文件excelFile
        File excelFile;
        if (isTemlate) {
            // 有模板情况复制模板文件到临时目录
            ResultData copyResult = copyTemplateFile(zipFile, targetTempFile);
            if (!copyResult.isSuccessful()) {
                return copyResult;
            }
            excelFile = new File(copyResult.getMessage());
        } else {
            String excelName = targetTempPath + File.separator + ExcelUtil.getBaseName(zipFile) + XLSX_SUFFIX_NAME;
            excelFile = new File(excelName);
        }
        ResultData initWorkbookResult;
        if (isTemlate) {
            initWorkbookResult = initWorkbook(excelFile);
            if (!initWorkbookResult.isSuccessful()) {
                return initWorkbookResult;
            }
        } else {
            sxssfWorkbook = new SXSSFWorkbook(FLUSH_DATA_SIZE);
        }
        CellStyleCreator styleCreator = new CellStyleCreator(sxssfWorkbook);
        List<SheetStyle> sheetStyles = excelStyle.getSheetStyle();
        if (isTemlate) {
            removeTemplateSheet();
        }
        XSSFFont xssfFont = (XSSFFont) sxssfWorkbook.getFontAt(0);
        String fontName = excelStyle.getDefaultFontName();
        if (fontName != null) {
            xssfFont.setFontName(fontName);
        }
        Short fontSize = excelStyle.getDefaultFontSize();
        if (fontSize != null) {
            xssfFont.setFontHeightInPoints(fontSize);
        }
        // 初始化fontMap
        Map<Integer, Font> fontMap = buildFont(styleCreator, excelStyle.getFontData());
        // 初始化cellStyleMap
        Map<String, CellStyle> cellStyleMap = buildCellStyleMap(sheetStyles, fontMap, styleCreator);
        csvSheetList.stream().forEach(csvSheet -> {
            String sheetName = csvSheet.getSheetName();
            if ((SummaryBookConst.COVER_SHEET_NAME.equals(sheetName) || SummaryBookConst.COVER_SHEET_NAME_CN.equals(
                sheetName)) && isTemlate) {
                // sumary 才特殊处理cover 非summary直接使用模板template的 cover页签
                if (fileType != null && TemplateType.of(fileType).isSummary()) {
                    createCoverSheet(csvSheet, styleCreator);
                }
            } else {
                Sheet sheet = sxssfWorkbook.getSheet(sheetName);
                if (sheet != null && isTemlate) {
                    sxssfWorkbook.removeSheetAt(sxssfWorkbook.getSheetIndex(sheet));
                }
                // 不存在模板直接按照原来数据还原
                sheet = sxssfWorkbook.createSheet(sheetName);
                SheetStyle targetSheetStyle = null;
                for (SheetStyle sheetStyle : sheetStyles) {
                    if (sheetStyle.getSheetName().equals(sheetName)) {
                        targetSheetStyle = sheetStyle;
                        break;
                    }
                }
                Set<Integer> addedCommentRows = Sets.newHashSet();
                try (CloseableIterable iterator = csvSheet.iterator()) {
                    int rowIndex = 0;
                    while (iterator.hasNext()) {
                        List<String> next = iterator.next();
                        Row rowData = sheet.createRow(rowIndex);
                        boolean isEmptyRow = isEmptyRow(next);
                        for (int column = 0; column < next.size(); column++) {
                            Cell cell = rowData.createCell(column);
                            if (!isEmptyRow) {
                                cell.setCellValue(Strings.nullToEmpty(next.get(column)));
                            }
                            CellStyle cellStyle = getCellStyle(rowIndex, column, targetSheetStyle, styleCreator, cellStyleMap);
                            if (cellStyle != null) {
                                cell.setCellStyle(cellStyle);
                            }
                            formatDate(cell);
                            // 设置Cell级样式
                            CellStyleFactory cellStyleFactory = new CellStyleFactory();
                            cellStyleFactory.setCellLevelStyle(cell, rowIndex, (short) column, targetSheetStyle,
                                cellStyleMap, styleCreator);
                            // 设置Column级样式
                            ColumnStyleFactory columnStyleFactory = new ColumnStyleFactory();
                            columnStyleFactory.setColumnLevelStyle(sheet, (short) column, targetSheetStyle);
                        }
                        // 设置Row级样式
                        RowStyleFactory rowStyleFactory = new RowStyleFactory();
                        rowStyleFactory.setRowLevelStyle(rowData, rowIndex, targetSheetStyle);
                        // SXSSF模式的Sheet会每隔FLUSH_DATA_SIZE行吧内存数据刷到xml中，所有这里需要在刷之前将行的批注写进去
                        if ((rowIndex + 1) % FLUSH_DATA_SIZE == 0) {
                            setComment(targetSheetStyle, sheet, fontMap, addedCommentRows);
                        }
                        rowIndex++;
                    }
                } catch (Exception e) {
                    LOGGER.error("save: Obtain iterator failed.", e);
                }
                // 设置批注
                setComment(targetSheetStyle, sheet, fontMap, addedCommentRows);
                // 设置Sheet级样式
                SheetStyleFactory sheetStyleFactory = new SheetStyleFactory();
                sheetStyleFactory.setSheetLevelStyle(sheet, targetSheetStyle);
                // 设置Workbook级样式
                WorkbookStyleFactory workbookStyleFactory = new WorkbookStyleFactory();
                workbookStyleFactory.setWorkbookLevelStyle(sxssfWorkbook, excelStyle);
            }
        });

        try (OutputStream fileOutStream = new FileOutputStream(excelFile)) {
            sxssfWorkbook.write(fileOutStream);
        } catch (FileNotFoundException e) {
            LOGGER.error("The file {} is not found", excelFile.getName());
            FileUtil.deleteDirectoryQuietly(tempDir);
            return new ResultData(false, MessageUtil.getMessage("FILE_NOT_FOUND_EN", "FILE_NOT_FOUND_CN", language));
        } catch (IOException e) {
            LOGGER.error("Init workbook fail.", e);
            FileUtil.deleteDirectoryQuietly(tempDir);
            return new ResultData(false,
                MessageUtil.getMessage("EXCEL_GENERATED_FAILED_EN", "EXCEL_GENERATED_FAILED_CN", language));
        }

        ResultData injectResult = null;
        try {
            injectResult = new MacroInjection().injectMacro(excelFile.getCanonicalPath());

        } catch (IOException e) {
            LOGGER.error("Init workbook fail.", e);
            FileUtil.deleteDirectoryQuietly(tempDir);
            return new ResultData(false,
                MessageUtil.getMessage("EXCEL_GENERATED_FAILED_EN", "EXCEL_GENERATED_FAILED_CN", language));
        }
        if (injectResult.isSuccessful()) {
            String finalPath = zipFile.getParent() + File.separator + ExcelUtil.getBaseName(excelFile)
                + XLSM_SUFFIX_NAME;
            File outputFile = new File(excelFile.getParent(), ExcelUtil.getBaseName(excelFile) + XLSM_SUFFIX_NAME);
            if (outputFile.exists()) {
                try {
                    Files.copy(outputFile, new File(finalPath));
                } catch (IOException e) {
                    FileUtil.deleteDirectoryQuietly(tempDir);
                    return new ResultData(false,
                        MessageUtil.getMessage("EXCEL_GENERATED_FAILED_EN", "EXCEL_GENERATED_FAILED_CN", language));
                } finally {
                    FileUtil.deleteDirectoryQuietly(tempDir);
                }
            } else {
                LOGGER.error("Init workbook fail.");
                return new ResultData(false,
                    MessageUtil.getMessage("EXCEL_GENERATED_FAILED_EN", "EXCEL_GENERATED_FAILED_CN", language));
            }
        } else {
            String finalPath = zipFile.getParent() + File.separator + ExcelUtil.getBaseName(excelFile)
                + XLSX_SUFFIX_NAME;
            try {
                Files.copy(excelFile, new File(finalPath));
            } catch (IOException e) {
                LOGGER.error("Init workbook fail.");
                return new ResultData(false,
                    MessageUtil.getMessage("EXCEL_GENERATED_FAILED_EN", "EXCEL_GENERATED_FAILED_CN", language));
            } finally {
                FileUtil.deleteDirectoryQuietly(tempDir);
            }
        }
        return new ResultData(true,
            MessageUtil.getMessage("EXCEL_GENERATED_SUCCESS_EN", "EXCEL_GENERATED_SUCCESS_CN", language));
    }

    /**
     * Negotiated 表示控制器表格，
     * 由于模板和zip 页签不一致，删除zip页签对应模板页签再次写入批注创建批注poi存在BUG，
     * 控制器表格数据恢复前，模板删除除cover之外的所有页签
     */
    private void removeTemplateSheet() {
        if (fileType != null && TemplateType.of(fileType).isNegotiated()) {
            ArrayList<Sheet> objects = new ArrayList<>();
            sxssfWorkbook.forEach(sheet -> {
                if (!(SummaryBookConst.COVER_SHEET_NAME.equals(sheet.getSheetName())
                    && !SummaryBookConst.COVER_SHEET_NAME_CN.equals(sheet.getSheetName()))
                    && !SUMMARYRES.equalsIgnoreCase(sheet.getSheetName())) {
                    objects.add(sheet);
                }
            });
            objects.stream().forEach(o -> sxssfWorkbook.removeSheetAt(sxssfWorkbook.getSheetIndex(o)));
        }
    }

    private boolean isEmptyRow(List<String> row) {
        return !Iterables.tryFind(row, (data) -> {
            return !Strings.isNullOrEmpty(data);
        }).isPresent();
    }

    private void setComment(SheetStyle sheetStyle, Sheet sheet, Map<Integer, Font> fontMap,
        Set<Integer> addedCommentRows) {
        if (sheetStyle == null || sheetStyle.getCellLevelStyle() == null) {
            return;
        }
        List<CommentData> commentDataList = sheetStyle.getCellLevelStyle().getCommentData();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        Set<Integer> tmpRows = Sets.newHashSet();
        for (CommentData commentData : commentDataList) {
            int rowIndex = commentData.getRow();
            int colIndex = commentData.getColumn();
            if (addedCommentRows.contains(rowIndex)) {
                continue;
            }
            Cell cell = getCell(sheet, rowIndex, colIndex);
            if (cell == null) {
                LOGGER.info("Cell[{}:{}] not exist!", rowIndex, colIndex);
                continue;
            }
            String[] strArray = commentData.getComment().split("\n");
            String[] newStrArray = getNewCommentArray(strArray);
            int endColumn = getEndColumn(sheet, colIndex, newStrArray);
            Font font = fontMap.get(commentData.getFontIndex());
            short fontHeight = font.getFontHeightInPoints();
            int endRow = getEndRow(sheet, rowIndex, newStrArray, fontHeight);
            ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, colIndex, rowIndex, endColumn, endRow);
            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
            Comment comment = drawing.createCellComment(anchor);
            RichTextString richTextString = sheet.getWorkbook()
                .getCreationHelper()
                .createRichTextString(commentData.getComment());
            richTextString.applyFont(font);
            comment.setString(richTextString);
            cell.removeCellComment();
            cell.setCellComment(comment);
            tmpRows.add(rowIndex);
        }
        addedCommentRows.addAll(tmpRows);
    }

    private String[] getNewCommentArray(String[] strArray) {
        List<String> commentStrList = new ArrayList<>();
        for (int i = 0; i < strArray.length; i++) {
            // 50为中文字符个数限制，100为英文字符个数限制
            int maxCharCount = isContainDoubleByte(strArray[i]) ? 50 : 100;
            if (strArray[i].length() > maxCharCount) {
                Pattern regex = Pattern.compile("(.{1," + maxCharCount + "}(?:\\s|$))|(.{0," + maxCharCount + "})", Pattern.DOTALL);
                Matcher regexMatcher = regex.matcher(strArray[i]);
                while (regexMatcher.find()) {
                    commentStrList.add(regexMatcher.group());
                }
            } else {
                commentStrList.add(strArray[i]);
            }
        }
        return commentStrList.toArray(new String[commentStrList.size()]);
    }

    private boolean isContainDoubleByte(String str) {
        Pattern p = Pattern.compile("[^x00-xff]");
        Matcher m = p.matcher(str);
        if (m.find()) {
            return true;
        }
        return false;
    }

    private Cell getCell(Sheet sheet, int row, int column) {
        Row rowData = sheet.getRow(row);
        if (rowData == null) {
            return null;
        }
        return rowData.getCell(column);
    }

    private int getEndRow(Sheet sheet, int rowIndex, String[] strArray, short fontHeight) {
        int rowHeight = sheet.getRow(rowIndex).getHeight();
        int commentHeight = strArray.length * (fontHeight * FONT_HEIGHT_MODULUS_A + FONT_HEIGHT_MODULUS_B);
        int endRow = rowIndex;
        if (rowHeight > commentHeight) {
            return endRow + 1;
        } else {
            while (commentHeight > 0) {
                int initHeight;
                Row row = sheet.getRow(endRow);
                endRow++;
                if (row != null) {
                    if (row.getZeroHeight()) {
                        continue;
                    }
                    initHeight = row.getHeight();
                } else {
                    initHeight = sheet.getDefaultRowHeight();
                }
                commentHeight = commentHeight - initHeight;
            }
            return endRow + 1;
        }
    }

    private int getEndColumn(Sheet sheet, int colIndex, String[] strArray) {
        int maxLength = getMaxLength(strArray);
        int endColumn = colIndex;
        while (maxLength > 0) {
            int initWidth = sheet.getColumnWidth(endColumn);
            endColumn++;
            int columnWidth;
            if (initWidth == DEFAULT_COLUMN_WIDTH) {
                columnWidth = sheet.getDefaultColumnWidth();
            } else {
                columnWidth = (initWidth - COLUMN_WIDTH_MODULUS_B) / COLUMN_WIDTH_MODULUS_A;
            }
            maxLength = maxLength - columnWidth;
        }
        return endColumn + 1;
    }

    private int getMaxLength(String[] strArray) {
        int maxLength = 0;
        for (int i = 0; i < strArray.length; i++) {
            Pattern p = Pattern.compile("[^x00-xff]");
            Matcher m = p.matcher(strArray[i]);
            int count = 0;
            while (m.find()) {
                count++;
            }
            int length = count * CHINESE_CHAR_MODULUS + (strArray[i].length() - count);
            if (length > maxLength) {
                maxLength = length;
            }
        }
        return maxLength;
    }

    private Map<String, CellStyle> buildCellStyleMap(List<SheetStyle> sheetStyles, Map<Integer, Font> fontMap, CellStyleCreator styleCreator) {
        Map<String, CellStyle> cellStyleMap = new HashMap<>();
        for (SheetStyle sheetStyle : sheetStyles) {
            CellLevelStyle cellLevelStyle = sheetStyle.getCellLevelStyle();
            if (cellLevelStyle == null) {
                continue;
            }
            List<StyleData> styleDatas = cellLevelStyle.getStyleData();
            for (StyleData styleData : styleDatas) {
                int cellStyleIndex = styleData.getCellStyleIndex();
                int fontIndex = styleData.getFontIndex();
                String styleKey = cellStyleIndex + "_" + fontIndex;
                CellStyle cellStyle = cellStyleMap.get(styleKey);
                if (cellStyle == null) {
                    cellStyle = buildCellStyle(cellStyleIndex, styleCreator, excelStyle.getCellStyleData());
                    cellStyle.setFont(fontMap.get(fontIndex));
                    cellStyleMap.put(styleKey, cellStyle);
                }
            }
            List<HyperlinkData> hyperlinkDatas = cellLevelStyle.getHyperlinkData();
            for (HyperlinkData hyperlinkData : hyperlinkDatas) {
                int cellStyleIndex = hyperlinkData.getCellStyleIndex();
                int fontIndex = hyperlinkData.getFontIndex();
                String styleKey = cellStyleIndex + "_" + fontIndex;
                CellStyle cellStyle = cellStyleMap.get(styleKey);
                if (cellStyle == null) {
                    cellStyle = buildCellStyle(cellStyleIndex, styleCreator, excelStyle.getCellStyleData());
                    cellStyle.setFont(fontMap.get(fontIndex));
                    cellStyleMap.put(styleKey, cellStyle);
                }
            }
        }
        return cellStyleMap;
    }

    private CellStyle getCellStyle(int rowIndex, int colIndex, SheetStyle sheetStyle, CellStyleCreator styleCreator,
        Map<String, CellStyle> cellStyleMap) {
        if (sheetStyle == null || sheetStyle.getCellLevelStyle() == null) {
            return null;
        }
        List<StyleData> styleDatas = sheetStyle.getCellLevelStyle().getStyleData();
        for (StyleData styleData : styleDatas) {
            CellRange cellRange = styleData.getCellRange();
            if (isCellInCellRange(rowIndex, colIndex, cellRange)) {
                CellStyle cellStyle;
                CellStyleType cellStyleType = styleData.getCellStyleType();
                if (cellStyleType != null && cellStyleType != CellStyleType.CUSTOMIZED) {
                    cellStyle = styleCreator.getStyleByType(styleData.getCellStyleType());
                } else {
                    int cellStyleIndex = styleData.getCellStyleIndex();
                    int fontIndex = styleData.getFontIndex();
                    String styleKey = cellStyleIndex + "_" + fontIndex;
                    cellStyle = cellStyleMap.get(styleKey);
                }
                return cellStyle;
            }
        }
        return null;
    }

    private void formatDate(Cell cell) {
        if (ExcelConstants.D_M_YY.equals(cell.getCellStyle().getDataFormatString()) && cell.getStringCellValue()
            .matches(ExcelConstants.DATE_REGEX)) {
            SimpleDateFormat originFormat = new SimpleDateFormat(ExcelConstants.DD_MM_YY);
            try {
                Date date = originFormat.parse(cell.getStringCellValue());
                SimpleDateFormat endFormat = new SimpleDateFormat(ExcelConstants.YYYY_MM_DD);
                cell.setCellValue(endFormat.format(date));
            } catch (ParseException e) {
                LOGGER.error("Parse date is error, date: {}", cell.getStringCellValue());
            }
        }
    }

    private boolean isCellInCellRange(int row, int column, CellRange cellRange) {
        return (cellRange.getRowBegin() <= row && row <= cellRange.getRowEnd()) && (cellRange.getColBegin() <= column
            && column <= cellRange.getColEnd());
    }

    private CellStyle buildCellStyle(int cellStyleIndex, CellStyleCreator styleCreator,
        List<CellStyleData> cellStyleDatas) {
        if (cellStyleDatas == null) {
            return styleCreator.getEmptyCellStyle();
        }
        for (CellStyleData cellStyleData : cellStyleDatas) {
            if (cellStyleData == null || cellStyleIndex != cellStyleData.getIndex()) {
                continue;
            }
            XSSFCellStyle cellStyle = (XSSFCellStyle) styleCreator.getEmptyCellStyle();
            // 设置自动换行
            if (cellStyleData.isWrapText() != null) {
                cellStyle.setWrapText(cellStyleData.isWrapText());
            }
            String rgbColor = cellStyleData.getBackgroundColor();
            if (rgbColor != null) {
                Color color = ExcelUtil.getRGBColor(rgbColor);
                if (color != null) {
                    XSSFColor xssfColor = new XSSFColor(color, new DefaultIndexedColorMap());
                    cellStyle.setFillForegroundColor(xssfColor);
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                } else {
                    // 兼容老的背景色设置，64和0为黑色，不设置
                    if (ExcelUtil.checkRgbColor(rgbColor)) {
                        cellStyle.setFillForegroundColor(Short.parseShort(rgbColor));
                        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    }
                }
            }
            buildStyle(styleCreator, cellStyleData, cellStyle);
            return cellStyle;
        }
        return styleCreator.getEmptyCellStyle();
    }

    private void buildStyle(CellStyleCreator styleCreator, CellStyleData cellStyleData, CellStyle cellStyle) {
        // 设置数据格式化样式
        if (!Strings.isNullOrEmpty(cellStyleData.getDataFormatString())) {
            cellStyle.setDataFormat(styleCreator.getDataFormat().getFormat(cellStyleData.getDataFormatString()));
        }
        // 设置水平对齐（居左，居中，居右）
        if (!Strings.isNullOrEmpty(cellStyleData.getAlignment())) {
            cellStyle.setAlignment(HorizontalAlignment.valueOf(cellStyleData.getAlignment()));
        }
        // 设置垂直对其
        if (!Strings.isNullOrEmpty(cellStyleData.getVerticalAlignment())) {
            cellStyle.setVerticalAlignment(
                VerticalAlignment.valueOf(VerticalAlignment.class, cellStyleData.getVerticalAlignment()));
        }
        // 设置下边框
        if (!Strings.isNullOrEmpty(cellStyleData.getBorderBottom())) {
            cellStyle.setBorderBottom(BorderStyle.valueOf(cellStyleData.getBorderBottom()));
        }
        // 设置左边框
        if (!Strings.isNullOrEmpty(cellStyleData.getBorderLeft())) {
            cellStyle.setBorderLeft(BorderStyle.valueOf(cellStyleData.getBorderLeft()));
        }
        // 设置右边框
        if (!Strings.isNullOrEmpty(cellStyleData.getBorderRight())) {
            cellStyle.setBorderRight(BorderStyle.valueOf(cellStyleData.getBorderRight()));
        }
        // 设置上边框
        if (!Strings.isNullOrEmpty(cellStyleData.getBorderTop())) {
            cellStyle.setBorderTop(BorderStyle.valueOf(cellStyleData.getBorderTop()));
        }
        // 设置单元格编辑锁
        if (cellStyleData.isLocked() != null) {
            cellStyle.setLocked(cellStyleData.isLocked());
        }
    }

    private Map<Integer, Font> buildFont(CellStyleCreator styleCreator, List<FontData> fontDatas) {
        Map<Integer, Font> map = new HashMap<>();
        if (fontDatas == null) {
            return map;
        }
        for (FontData fontData : fontDatas) {
            if (fontData == null) {
                continue;
            }
            XSSFFont font = (XSSFFont) styleCreator.getEmptyFont();
            if (fontData.getFontSize() != null) {
                font.setFontHeightInPoints(fontData.getFontSize());
            }
            if (fontData.isBold() != null) {
                font.setBold(fontData.isBold());
            }
            if (fontData.getColor() != null) {
                Color color = ExcelUtil.getRGBColor(fontData.getColor());
                if (color != null) {
                    font.setColor(new XSSFColor(color, new DefaultIndexedColorMap()));
                }
            }
            String fontName = fontData.getFontName();
            if (!Strings.isNullOrEmpty(fontName)) {
                font.setFontName(fontName);
            }
            Byte underLine = fontData.getUnderline();
            if (underLine != null) {
                font.setUnderline(underLine);
            }
            map.put(fontData.getIndex(), font);
        }
        return map;
    }

    private void createCoverSheet(CsvSheet csvSheet, CellStyleCreator styleCreator) {
        Map<String, String> coverMap = new TreeMap<>();
        String title = "";
        try (CloseableIterable iterator = csvSheet.iterator()) {
            int rowIndex = 0;

            while (iterator.hasNext()) {
                if (rowIndex > MAX_TITLE_ROW) {
                    break;
                }
                if (rowIndex == TITLE_ROW) {
                    title = Strings.nullToEmpty(iterator.next().get(TITLE_START_COL));
                    rowIndex++;
                    continue;
                }
                if (rowIndex < COVER_START_ROW) {
                    iterator.next();
                    rowIndex++;
                    continue;
                }
                List<String> next = iterator.next();
                String key = Strings.nullToEmpty(next.get(TITLE_START_COL));
                //du dao na ting zhi issue
                String value = Strings.nullToEmpty(next.get(TITLE_VALUE_COL));
                coverMap.put(key, value);
                rowIndex++;
            }
        } catch (Exception e) {
            LOGGER.error("save: Obtain iterator failed.", e);
        }

        SummaryCoverSheet coverSheet = new SummaryCoverSheet(sxssfWorkbook, styleCreator);
        coverSheet.motifyCoverSheet(title, coverMap);
    }

    /**
     * 获取模板路径
     *
     * @param list Internal文件下的FileIdentification文件 中的key 和filetype集合
     * @return TemplatePath 模板路径
     */
    private String getTemplatePath(List<String> list) {
        String path;
        String finalTemplate = list.get(0);
        fileType = list.get(1);
        // 获取filecodeMapping 中的键值对
        Map<String, List<String>> configMap = ConfigUtil.getFileCodeConfig();
        Optional<Map.Entry<String, List<String>>> optional = configMap.entrySet()
            .stream()
            .filter(e -> e.getKey().equals(finalTemplate))
            .findAny();
        if (optional.isPresent()) {
            Map.Entry<String, List<String>> pathSet = optional.get();
            List<String> resultList = pathSet.getValue();
            if (resultList.isEmpty()) {
                return "";
            }
            path = resultList.get(1);
        } else {
            return "";
        }
        if (TemplateType.of(fileType).isControllerSummary()) {
            for (CsvSheet csvSheet : csvSheetList) {
                String sheetName = csvSheet.getSheetName();
                if (SummaryBookConst.COVER_SHEET_NAME.equals(sheetName) || SummaryBookConst.COVER_SHEET_NAME_CN.equals(
                    sheetName)) {
                    path = ConfigUtil.getFinalTemplatePath(path, fileType, csvSheet);
                    break;
                }
            }
        }
        if (path != null && !path.isEmpty()) {
            path = MACRO_TOOL_PATH + path;
        }

        return path;
    }
}