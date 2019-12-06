/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.workbook;

import com.app.excelcsvconverter.xmlmodel.styledata.ExcelStyle;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * 创建Workbook级样式的工厂
 *
 * @since 2019-07-22
 */
public class WorkbookStyleFactory {

    private List<WorkbookStyleOperator> workbookStyleOperators = new ArrayList<>();

    private List<WorkbookStyleOperator> getWorkbookStyleOperators() {
        ActiveSheetSetter activeSheetSetter = new ActiveSheetSetter();
        workbookStyleOperators.add(activeSheetSetter);
        SheetVisibilitySetter sheetVisibilitySetter = new SheetVisibilitySetter();
        workbookStyleOperators.add(sheetVisibilitySetter);
        SheetNamesSorter sheetNamesSorter = new SheetNamesSorter();
        workbookStyleOperators.add(sheetNamesSorter);
        return workbookStyleOperators;
    }

    public void setWorkbookLevelStyle(Workbook workbook, ExcelStyle excelStyle) {
        if (excelStyle == null) {
            return;
        }
        for (WorkbookStyleOperator workbookStyleOperator : getWorkbookStyleOperators()) {
            workbookStyleOperator.setStyle(workbook, excelStyle);
        }
    }
}