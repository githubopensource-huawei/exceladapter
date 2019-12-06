/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.sheet;

import com.app.excelcsvconverter.xmlmodel.styledata.CellRange;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * 设置合并单元格的区域
 *
 * @since 2019-07-22
 */
public class MergeRegionSetter implements SheetStyleOperator {

    @Override
    public void setStyle(Sheet sheet, SheetStyle sheetStyle) {
        List<CellRange> mergedRange = sheetStyle.getCellRange();
        for (CellRange cellRange : mergedRange) {
            if (needMerge(cellRange)) {
                CellRangeAddress cellRangeAddress = buildCellRangeAddress(cellRange);
                if (cellRangeAddress.getNumberOfCells() > 1) {
                    sheet.addMergedRegion(cellRangeAddress);
                }
            }
        }
    }

    /**
     * 判断是否需要合并
     *
     * @return
     */
    private boolean needMerge(CellRange cellRange) {
        return Math.abs(cellRange.getRowEnd() - cellRange.getRowBegin()) + Math.abs(cellRange.getColEnd() - cellRange.getColBegin()) > 0;
    }

    /**
     * 构建CellRangeAddress对象
     *
     * @return
     */
    private CellRangeAddress buildCellRangeAddress(CellRange cellRange) {
        return new CellRangeAddress(cellRange.getRowBegin(), cellRange.getRowEnd(), cellRange.getColBegin(), cellRange.getColEnd());
    }
}