package com.app.excelcsvconverter.style.cell;

import com.app.excelcsvconverter.style.CellStyleCreator;
import com.app.excelcsvconverter.xmlmodel.styledata.CellStyleType;
import com.app.excelcsvconverter.xmlmodel.styledata.HyperlinkData;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;
import com.google.common.base.Strings;

import java.util.List;
import java.util.Map;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 设置超链接
 *
 * @since 2019-07-22
 */
public class HyperlinkSetter implements CellStyleOperator {

    private static final Logger LOGGER = LoggerFactory.getLogger(HyperlinkSetter.class);

    @Override
    public void setStyle(Cell cell, int rowIndex, short colIndex, SheetStyle sheetStyle,
        Map<String, CellStyle> cellStyleMap, CellStyleCreator styleCreator) {
        if (sheetStyle.getCellLevelStyle() == null) {
            return;
        }
        List<HyperlinkData> hyperlinkDatas = sheetStyle.getCellLevelStyle().getHyperlinkData();
        for (HyperlinkData hyperlinkData : hyperlinkDatas) {
            if (hyperlinkData == null) {
                continue;
            }
            if (rowIndex == hyperlinkData.getRow() && colIndex == hyperlinkData.getColumn()) {
                if (Strings.isNullOrEmpty(cell.getStringCellValue())) {
                    LOGGER.info("Cell[row:%s column:%s] not exist.", hyperlinkData.getRow(), hyperlinkData.getColumn());
                } else {
                    Workbook workbook = cell.getRow().getSheet().getWorkbook();
                    Hyperlink link = workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
                    link.setAddress(buildAddress(hyperlinkData));
                    cell.setHyperlink(link);
                    CellStyleType cellStyleType = hyperlinkData.getCellStyleType();
                    CellStyle hyperlinkStyle;
                    if (cellStyleType != null && cellStyleType != CellStyleType.CUSTOMIZED) {
                        hyperlinkStyle = styleCreator.getStyleByType(hyperlinkData.getCellStyleType());
                        cell.setCellStyle(hyperlinkStyle);
                    } else {
                        int cellStyleIndex = hyperlinkData.getCellStyleIndex();
                        int fontIndex = hyperlinkData.getFontIndex();
                        String styleKey = cellStyleIndex + "_" + fontIndex;
                        hyperlinkStyle = cellStyleMap.get(styleKey);
                        if (hyperlinkStyle != null) {
                            cell.setCellStyle(hyperlinkStyle);
                        }
                    }
                }
                return;
            }
        }
    }

    private String buildAddress(HyperlinkData hyperlinkData) {
        return String.format("'%s'!R%dC%d", hyperlinkData.getLinkSheetName(), hyperlinkData.getLinkRow(),
            hyperlinkData.getLinkColumn());
    }
}