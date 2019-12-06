/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.sheet;

import com.app.excelcsvconverter.util.ExcelUtil;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;
import com.google.common.base.Strings;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.awt.Color;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

/**
 * 设置页签导航栏的颜色
 *
 * @since 2019-07-22
 */
public class TabColorSetter implements SheetStyleOperator {
    private Map<String, Color> colorMap = new HashMap<>();

    public TabColorSetter() {
        colorMap.put("RED", Color.RED);
        colorMap.put("GREEN", Color.GREEN);
        colorMap.put("BLACK", Color.BLACK);
        colorMap.put("BLUE", Color.BLUE);
        colorMap.put("YELLOW", Color.YELLOW);
        colorMap.put("GRAY", Color.GRAY);
        colorMap.put("PINK", Color.PINK);
        colorMap.put("ORANGE", Color.ORANGE);
        colorMap.put("LIGHT_GRAY", Color.LIGHT_GRAY);
        colorMap.put("DARK_GRAY", Color.DARK_GRAY);
        colorMap.put("MAGENTA", Color.MAGENTA);
        colorMap.put("CYAN", Color.CYAN);
    }

    @Override
    public void setStyle(Sheet sheet, SheetStyle sheetStyle) {
        String tabColor = sheetStyle.getTabColor();
        if (Strings.isNullOrEmpty(tabColor)) {
            return;
        }

        Color color = colorMap.get(tabColor.toUpperCase(Locale.ENGLISH));
        if (color == null) {
            color = ExcelUtil.getRGBColor(tabColor);
        }
        if (color == null) {
            return;
        }
        if (sheet instanceof SXSSFSheet) {
            SXSSFSheet sxssfSheet = (SXSSFSheet) sheet;
            sxssfSheet.setTabColor(new XSSFColor(color, new DefaultIndexedColorMap()));
        } else if (sheet instanceof XSSFSheet) {
            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            xssfSheet.setTabColor(new XSSFColor(color, new DefaultIndexedColorMap()));
        }
    }
}