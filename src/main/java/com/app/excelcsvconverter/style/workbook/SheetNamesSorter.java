/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.workbook;

import com.app.excelcsvconverter.xmlmodel.styledata.ExcelStyle;
import com.google.common.base.Strings;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

public class SheetNamesSorter implements WorkbookStyleOperator {

    private static final String SEPARATOR = ",";

    @Override
    public void setStyle(Workbook workbook, ExcelStyle excelStyle) {
        String sheetName = excelStyle.getSortedSheetNames();
        if (Strings.isNullOrEmpty(sheetName)) {
            return;
        }
        String[] tmp = sheetName.split(SEPARATOR);
        List<String> sheetNames = Arrays.asList(tmp);
        List<String> allSheets = new ArrayList<>();
        for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
            allSheets.add(workbook.getSheetAt(i).getSheetName());
        }
        allSheets.sort(Comparator.comparingInt(sheetNames::indexOf));
        for (int i = 0; i < allSheets.size(); ++i) {
            workbook.setSheetOrder(allSheets.get(i), i);
        }
    }
}