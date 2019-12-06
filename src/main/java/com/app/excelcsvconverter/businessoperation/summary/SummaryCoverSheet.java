/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.businessoperation.summary;

import static com.app.excelcsvconverter.consts.summary.SummaryConfig.CONFIG;

import com.app.excelcsvconverter.consts.summary.SummaryBookConst;
import com.app.excelcsvconverter.style.CellStyleCreator;
import com.google.common.base.Strings;
import com.microsoft.schemas.vml.impl.CTShapetypeImpl;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.XmlObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/**
 * Summary Cover页签处理
 *
 * @since 2019-08-08
 */
public class SummaryCoverSheet {

    public static final int HEAD_COL_INDEX = 1;

    public static final int VERSION_COL_INDEX = 3;

    public static final int TITLE_START_COL = 1;

    public static final int TITLE_VALUE_COL = 3;

    public static final int TITLE_END_COL = 4;

    public static final int MAX_TITLE_ROW = 6;

    public static final int COVER_START_ROW = 2;

    public static final int TITLE_ROW = 1;

    private static final Logger LOGGER = LoggerFactory.getLogger(SummaryCoverSheet.class);

    private XSSFWorkbook xssfWorkbook;

    private CellStyleCreator styleCreator;

    private Map<String, String> coverMap = new TreeMap<>();

    public SummaryCoverSheet(SXSSFWorkbook sxssfWorkbook, CellStyleCreator styleCreator) {
        this.xssfWorkbook = sxssfWorkbook.getXSSFWorkbook();
        this.styleCreator = styleCreator;
    }

    public void motifyCoverSheet(String title, Map<String, String> coverMap) {
        this.coverMap = coverMap;

        if (coverMap.containsKey(CONFIG.getConfig("COVER_SCENARIO_EN"))) {
            addScenario(coverMap.get(CONFIG.getConfig("COVER_SCENARIO_EN")));
        } else if (coverMap.containsKey(CONFIG.getConfig("COVER_SCENARIO_CN"))) {
            addScenario(coverMap.get(CONFIG.getConfig("COVER_SCENARIO_CN")));
        }

        boolean isOldFormat = coverMap.containsKey(CONFIG.getConfig("OldGsmVersion_CN")) || coverMap.containsKey(
            CONFIG.getConfig("OldGsmVersion_EN"));
        modifySummaryVersion(isOldFormat);
        modifyCoverTitle(title);
        modifyCoverIssueDate();
    }

    public boolean modifyCoverTitle(String typeInfo) {
        if (Strings.isNullOrEmpty(typeInfo)) {
            return false;
        }
        Sheet coverSheet = xssfWorkbook.getSheet(SummaryBookConst.COVER_SHEET_NAME);
        if (coverSheet == null) {
            coverSheet = xssfWorkbook.getSheet(SummaryBookConst.COVER_SHEET_NAME_CN);
        }

        if (coverSheet == null) {
            LOGGER.info("The cover sheet is not found.");
            return false;
        }

        for (int i = 0; i <= MAX_TITLE_ROW; i++) {
            Cell cell = getCell(coverSheet, i, 1);
            if (cell != null && "Cover Title".equalsIgnoreCase(cell.getStringCellValue())) {
                cell.setCellValue(typeInfo);
                updateCellValue(coverSheet, cell, typeInfo);
                break;
            }
        }

        return true;
    }

    public boolean modifyCoverIssueDate() {
        Sheet coverSheet = xssfWorkbook.getSheet(SummaryBookConst.COVER_SHEET_NAME);
        if (coverSheet == null) {
            coverSheet = xssfWorkbook.getSheet(SummaryBookConst.COVER_SHEET_NAME_CN);
        }
        if (coverSheet == null) {
            LOGGER.info("The cover sheet is not found.");
            return false;
        }
        for (int i = 0; i <= MAX_TITLE_ROW; i++) {
            Cell cell = getCell(coverSheet, i, 1);
            if (cell != null && (CONFIG.getConfig("ISSUE_DATE_CN").equalsIgnoreCase(cell.getStringCellValue()) || CONFIG
                .getConfig("ISSUE_DATE_EN")
                .equalsIgnoreCase(cell.getStringCellValue()))) {
                String issueDate = coverMap.get(cell.getStringCellValue());
                Cell cellDate = getCell(coverSheet, i, 3);
                cellDate.setCellValue(issueDate);
                updateCellValue(coverSheet, cellDate, issueDate);
                break;
            }
        }

        return true;
    }

    public boolean modifySummaryVersion(boolean isOldFormat) {
        int startRow = 2;
        Sheet coverSheet = xssfWorkbook.getSheet(SummaryBookConst.COVER_SHEET_NAME);
        String issueDate;
        String metrology;
        String gsmControllerMode;
        String umtsControllerMode;
        String scenario;
        boolean isCn = false;
        if (coverSheet == null) {
            coverSheet = xssfWorkbook.getSheet(SummaryBookConst.COVER_SHEET_NAME_CN);
            issueDate = CONFIG.getConfig("ISSUE_DATE_CN");
            metrology = CONFIG.getConfig("METROLOGY_CN");
            gsmControllerMode = CONFIG.getConfig("COVER_GSM_MODE_CN");
            umtsControllerMode = CONFIG.getConfig("COVER_UMTS_MODE_CN");
            scenario = CONFIG.getConfig("COVER_SCENARIO_CN");
            isCn = true;
        } else {
            issueDate = CONFIG.getConfig("ISSUE_DATE_EN");
            metrology = CONFIG.getConfig("METROLOGY_EN");
            gsmControllerMode = CONFIG.getConfig("COVER_GSM_MODE_EN");
            umtsControllerMode = CONFIG.getConfig("COVER_UMTS_MODE_EN");
            scenario = CONFIG.getConfig("COVER_SCENARIO_EN");
        }

        if (coverSheet == null) {
            LOGGER.info("The cover sheet is not found.");
            return false;
        }

        int count = 0;
        boolean isFound = false;
        for (int i = startRow; i <= MAX_TITLE_ROW; i++) {
            Cell cell = getCell(coverSheet, i, 1);
            if (cell != null && (issueDate.equals(cell.getStringCellValue()) || metrology.equals(
                cell.getStringCellValue()) || gsmControllerMode.equals(cell.getStringCellValue().trim())
                || umtsControllerMode.equals(cell.getStringCellValue().trim()) || scenario.equals(
                cell.getStringCellValue().trim()))) {
                isFound = true;
                break;
            }
            count++;
        }

        if (!isFound && isOldFormat) {
            //如果是GBTS建站，由于其Cover页中英文都叫Cover，所以需要特殊处理一下，设置为中文之后，重新获取版本号行数
            issueDate = CONFIG.getConfig("ISSUE_DATE_CN");
            metrology = CONFIG.getConfig("METROLOGY_CN");
            gsmControllerMode = CONFIG.getConfig("COVER_GSM_MODE_CN");
            umtsControllerMode = CONFIG.getConfig("COVER_UMTS_MODE_CN");
            scenario = CONFIG.getConfig("COVER_SCENARIO_CN");
            isCn = true;
            // 模板中version的个数
            count = 0;  //count指的是基础表格中版本的个数
            for (int i = startRow; i <= MAX_TITLE_ROW; i++) {
                Cell cell = getCell(coverSheet, i, 1);
                if (cell != null && (issueDate.equals(cell.getStringCellValue()) || metrology.equals(
                    cell.getStringCellValue()) || gsmControllerMode.equals(cell.getStringCellValue().trim())
                    || umtsControllerMode.equals(cell.getStringCellValue().trim()) || scenario.equals(
                    cell.getStringCellValue().trim()))) {
                    isFound = true;
                    break;
                }
                count++;
            }
        }
        int versionSize = getVerionSize();

        if (!isFound || count > versionSize) { //需要添加的version个数    >的作用是保证是扩行操作
            LOGGER.info("The format of the cover sheet has error.");
            return false;
        }

        initCoverHeadVerStyle(coverSheet, startRow);

        int shiftRowNum = versionSize - count;
        List<CellRangeAddress> mergedRegions = buildShiftRowsCellRange(startRow, coverSheet.getMergedRegions(),
            shiftRowNum);
        if (count < versionSize) {
            coverSheet.shiftRows(startRow, coverSheet.getLastRowNum(), shiftRowNum, true, false);
            solveShiftRowsMethodError(coverSheet, startRow, coverSheet.getLastRowNum(), shiftRowNum);
        }

        for (int i = startRow; i <= startRow + versionSize - count - 1; i++) {
            addTitleRow(coverSheet, i);
        }

        addMergedRegion(coverSheet, mergedRegions);

        if (isOldFormat) {
            initOldGsmVersion(coverSheet, isCn, startRow);
            return true;
        }

        for (Map.Entry<String, String> value : coverMap.entrySet()) {
            if (isVersion(value.getKey())) {
                initVersion(coverSheet, startRow, value.getKey(), value.getValue());
                startRow++;
            }
        }

        return true;
    }

    /**
     * 规避poi 4.0.1版本shiftRows方法问题，poi计划4.1.1版本解决此问题
     * bug id:57423
     */
    private void solveShiftRowsMethodError(Sheet sheet, int firstRow, int lastRow, int numToShift) {
        int firstDstRow = firstRow + numToShift;
        int lastDstRow = lastRow + numToShift;
        for (int row = firstDstRow; row <= lastDstRow; ++row) {
            final XSSFRow xssfRow = (XSSFRow) sheet.getRow(row);
            if (xssfRow != null) {
                String msg = "Row[rownum=" + xssfRow.getRowNum()
                    + "] contains cell(s) included in a multi-cell array formula. "
                    + "You cannot change part of an array.";
                for (Cell c : xssfRow) {
                    ((XSSFCell) c).updateCellReferencesForShifting(msg);
                }
            }
        }
    }

    // 规避poi 4.0.1版本setCellValue方法无效
    private void updateCellValue(Sheet sheet, Cell cell, String value) {
        CellStyle cellStyle = cell.getCellStyle();
        Cell newCell = sheet.getRow(cell.getRowIndex()).createCell(cell.getColumnIndex());
        newCell.setCellValue(value);
        newCell.setCellStyle(cellStyle);
    }

    private int getVerionSize() {
        if (coverMap == null || coverMap.isEmpty()) {
            return 0;
        }
        int count = 0;
        Set<String> keys = coverMap.keySet();
        for (String key : keys) {
            if (key.contains(CONFIG.getConfig("Version_EN")) || key.contains("version") || key.contains(
                CONFIG.getConfig("Version_CN"))) {
                count++;
            }
        }
        return count;
    }

    private boolean isVersion(String title) {
        return title.contains(CONFIG.getConfig("Version_EN")) || title.contains("version") || title.contains(
            CONFIG.getConfig("Version_CN"));
    }

    private boolean initOldGsmVersion(Sheet coverSheet, boolean isCn, int startRow) {

        String title = isCn ? CONFIG.getConfig("OldGsmVersion_CN") : CONFIG.getConfig("OldGsmVersion_EN");
        String dspVersion = isCn
            ? coverMap.get(CONFIG.getConfig("OldGsmVersion_CN"))
            : coverMap.get(CONFIG.getConfig("OldGsmVersion_EN"));
        return initVersion(coverSheet, startRow, title, dspVersion);
    }

    private boolean initVersion(Sheet coverSheet, int startRow, String title, String version) {
        boolean result = false;
        if (!Strings.isNullOrEmpty(version)) {
            Cell cell = getCell(coverSheet, startRow, 1);
            if (cell != null) {
                cell.setCellValue(title);
                updateCellValue(coverSheet, cell, title);
                cell = getCell(coverSheet, startRow, 3);
                Cell mergeCell = getCell(coverSheet, startRow, 4);
                if (cell != null && mergeCell != null) {
                    cell.setCellValue(version);
                    updateCellValue(coverSheet, cell, version);
                    if (version.contains(";")) {
                        List<CommentData> comments = new ArrayList<>();
                        String replace = version.replace(";", "\n");
                        comments.add(new CommentData(startRow, (short) 3, replace));
                        addComments(coverSheet.getSheetName(), comments);
                    }
                }
                result = true;
            }
        }
        return result;
    }

    private void repairArrowComment(XSSFSheet sheet) {
        try {
            Class<?> xssfSheetClass = Class.forName("org.apache.poi.xssf.usermodel.XSSFSheet");
            Method vmlMethod = xssfSheetClass.getDeclaredMethod("getVMLDrawing", Boolean.TYPE);
            vmlMethod.setAccessible(true);
            XSSFVMLDrawing drawing = (XSSFVMLDrawing) vmlMethod.invoke(sheet, new Object[] {Boolean.TRUE});
            Class<?> drawClass = Class.forName("org.apache.poi.xssf.usermodel.XSSFVMLDrawing");
            Field field = drawClass.getDeclaredField("_shapeTypeId");
            field.setAccessible(true);
            field.set(drawing, "_x0000_t202");
            Method method = drawClass.getDeclaredMethod("getItems");
            method.setAccessible(true);
            List<XmlObject> items = (List) method.invoke(drawing, new Object[0]);
            items.stream()
                .filter((item) -> item instanceof CTShapetypeImpl && "_xssf_cell_comment".equals(
                    ((CTShapetypeImpl) item).getId()))
                .forEach((shape) -> ((CTShapetypeImpl) shape).setId("_x0000_t202"));
        } catch (Exception var9) {
            LOGGER.error("Arrows may appear again! I feel so sorry...", var9);
        }

    }

    public void addComments(String sheetName, List<CommentData> comments) {
        XSSFSheet sheet = xssfWorkbook.getSheet(sheetName);
        if (sheet == null) {
            LOGGER.info("{} not exist.", sheetName);
        } else {
            this.repairArrowComment(sheet);
            XSSFDrawing patriarch = sheet.createDrawingPatriarch();
            comments.forEach((commentData) -> {
                int row = commentData.getRow();
                short column = commentData.getColumn();
                Cell cell = this.getCell(sheet, row, column);
                if (cell == null) {
                    LOGGER.info("Cell[{}:{}] not exist!", row, column);
                } else {
                    cell.removeCellComment();
                    XSSFComment comment = createComment(sheet, patriarch, commentData.getMessage(), "CME", row, column);
                    cell.setCellComment(comment);
                }

            });
        }
    }

    protected XSSFComment createComment(Sheet sheet, XSSFDrawing patriarch, String str, String author, int row,
        short column) {
        String[] strArray = str.split("\n");
        int endColumn = getEndColumn(sheet, row, column, strArray);
        XSSFClientAnchor anchor = patriarch.createAnchor(0, 0, 0, 0, column + 1, row, endColumn + 1,
            row + strArray.length + 1);
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        XSSFComment comment = patriarch.createCellComment(anchor);
        RichTextString richTextString = new XSSFRichTextString(str);
        richTextString.applyFont(this.styleCreator.getCommentFont());
        comment.setString(richTextString);
        comment.setAuthor(author);
        return comment;
    }

    protected int getEndColumn(Sheet sheet, int row, short column, String[] strArray) {
        int maxLength = 0;
        String[] var6 = strArray;
        int var7 = strArray.length;

        for (int var8 = 0; var8 < var7; ++var8) {
            String line = var6[var8];
            maxLength = line.length() > maxLength ? line.length() : maxLength;
        }

        int endColumn;
        Cell cell;
        for (
            endColumn = column;
            maxLength > 0; maxLength = cell != null && !Strings.isNullOrEmpty(cell.getStringCellValue()) ? maxLength
            - cell.getStringCellValue().length() : maxLength - 5) {
            ++endColumn;
            cell = this.getCell(sheet, row, endColumn);
        }
        return endColumn;
    }

    private void initCoverHeadVerStyle(Sheet coverSheet, int startRow) {
        try {
            CellStyle existCoverHeadStyle = getCell(coverSheet, startRow, HEAD_COL_INDEX).getCellStyle();
            CellStyle existCoverVerStyle = getCell(coverSheet, startRow, VERSION_COL_INDEX).getCellStyle();
            styleCreator.setSummaryCoverHeadStyle(existCoverHeadStyle);
            styleCreator.setSummaryCoverVersionStyle(existCoverVerStyle);
        } catch (Exception e) {
            LOGGER.error("init Cover Head and Version Style failed!", e);
        }
    }

    public boolean addScenario(String scenario) {
        if (Strings.isNullOrEmpty(scenario)) {
            return true;
        }

        Sheet coverSheet = xssfWorkbook.getSheet(SummaryBookConst.COVER_SHEET_NAME);
        boolean isCn = false;
        if (coverSheet == null) {
            isCn = true;
            coverSheet = xssfWorkbook.getSheet(SummaryBookConst.COVER_SHEET_NAME_CN);
            if (coverSheet == null) {
                LOGGER.error("The cover sheet is not found.");
                return false;
            }
        }
        String issueDate = isCn ? CONFIG.getConfig("ISSUE_DATE_CN") : CONFIG.getConfig("ISSUE_DATE_EN");
        String metrology = isCn ? CONFIG.getConfig("METROLOGY_CN") : CONFIG.getConfig("METROLOGY_EN");
        String scenarioDesc = isCn ? CONFIG.getConfig("COVER_SCENARIO_CN") : CONFIG.getConfig("COVER_SCENARIO_EN");

        for (int i = COVER_START_ROW; i <= MAX_TITLE_ROW; i++) {
            Cell cell = getCell(coverSheet, i, TITLE_START_COL);
            if (cell != null && scenarioDesc.equals(cell.getStringCellValue())) {
                return true;
            }
            if (cell != null) {
                if ((issueDate.equals(cell.getStringCellValue()) || metrology.equals(cell.getStringCellValue()))) {
                    int startRow = i - 1;
                    List<CellRangeAddress> mergedRegions = buildShiftRowsCellRange(startRow,
                        coverSheet.getMergedRegions(), 1);
                    coverSheet.shiftRows(startRow, coverSheet.getLastRowNum(), 1, true, false);
                    addTitleRow(coverSheet, startRow);
                    addMergedRegion(coverSheet, mergedRegions);
                    Cell startCell = getCell(coverSheet, i, TITLE_START_COL);
                    startCell.setCellValue(scenarioDesc);
                    updateCellValue(coverSheet, startCell, scenarioDesc);
                    Cell endCell = getCell(coverSheet, i, TITLE_VALUE_COL);
                    endCell.setCellValue(scenario);
                    updateCellValue(coverSheet, endCell, scenario);
                    break;
                }
            }
        }
        return true;
    }

    public Cell getCell(Sheet sheet, int row, int column) {
        XSSFSheet xssfSheet = this.xssfWorkbook.getSheet(sheet.getSheetName());
        if (xssfSheet == null) {
            return null;
        } else {
            Row rowData = xssfSheet.getRow(row);
            return rowData == null ? null : rowData.getCell(column);
        }
    }

    /**
     * POI3.15，扩行操作对03/07的行为不一致，03英文版的先扩所有行再清空，而中午版先扩所涉及行再清空，导致原MergeCell有变动，
     * 而07先清空再扩行，原MergeCell不会改变;
     * 避免特殊处理，直接记录MergeCell，自行做扩行操作。
     */
    private List<CellRangeAddress> buildShiftRowsCellRange(int startRow, List<CellRangeAddress> regions,
        int shiftRowNum) {
        List<CellRangeAddress> mergedRegions = new ArrayList<>();
        for (CellRangeAddress region : regions) {
            mergedRegions.add(new CellRangeAddress(region.getFirstRow(), region.getLastRow(), region.getFirstColumn(),
                region.getLastColumn()));
        }
        for (CellRangeAddress range : mergedRegions) {
            if (range.getFirstRow() > startRow) {
                range.setFirstRow(range.getFirstRow() + shiftRowNum);
                range.setLastRow(range.getLastRow() + shiftRowNum);
            }
        }
        return mergedRegions;
    }

    private void addTitleRow(Sheet coverSheet, int startRow) {
        Row row = coverSheet.createRow(startRow);
        for (int i = TITLE_START_COL; i <= TITLE_END_COL; i++) {
            Cell cell = row.createCell(i);
            if (i == 1 || i == 2) {
                cell.setCellStyle(styleCreator.getSummaryCoverHeadStyle());
            } else if (i == 3 || i == 4) {
                cell.setCellStyle(styleCreator.getSummaryCoverVersionStyle());
            }
        }
        coverSheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 1, 2));
        coverSheet.addMergedRegion(new CellRangeAddress(startRow, startRow, 3, 4));
    }

    private void addMergedRegion(Sheet coverSheet, List<CellRangeAddress> mergedRegions) {
        List<CellRangeAddress> newMergedRegions = coverSheet.getMergedRegions();
        for (CellRangeAddress range : mergedRegions) {
            if (!newMergedRegions.contains(range)) {
                if (range.getNumberOfCells() > 1) {
                    coverSheet.addMergedRegion(range);
                }
            }
        }
    }
}
