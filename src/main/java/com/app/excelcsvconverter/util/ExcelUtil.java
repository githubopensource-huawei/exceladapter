/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.util;

import com.app.excelcsvconverter.parser.ExcelFormat;

import org.apache.poi.poifs.filesystem.FileMagic;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Color;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * 功能描述: Excel工具类
 *
 * @since 2019-04-30
 */
public class ExcelUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * Check format excel format.
     *
     * @param file the file
     * @return the excel format
     */
    public static ExcelFormat checkFormat(File file) {
        try (InputStream in = new FileInputStream(file); BufferedInputStream bufIn = new BufferedInputStream(in, 8)) {
            if (FileMagic.valueOf(bufIn) == FileMagic.OLE2) {
                return ExcelFormat.OLE2;
            }
            if (FileMagic.valueOf(bufIn) == FileMagic.OOXML) {
                return ExcelFormat.OOXML;
            }
        } catch (IOException ex) {
            LOGGER.error("EXCEPTION occurred when check excel format:{}", file.getName(), ex);
        }
        return ExcelFormat.NA;
    }

    public static Color getRGBColor(String tabColor) {
        String[] rgb = tabColor.split(ExcelConstants.SPLIT_REGEX);
        if (rgb.length != 3) {
            return null;
        }
        try {
            int r = Integer.parseInt(rgb[0]);
            int g = Integer.parseInt(rgb[1]);
            int b = Integer.parseInt(rgb[2]);
            Color color = new Color(r, g, b);
            return color;
        } catch (NumberFormatException e) {
            LOGGER.error("rgb value is wrong", e);
            return null;
        }
    }

    public static boolean checkRgbColor(String rgbColor) {
        return !rgbColor.contains(",") && !"64".equals(rgbColor) && !"0".equals(rgbColor) && rgbColor != null
            && !rgbColor.isEmpty();
    }

    public static boolean isCoverSheet(String sheetName) {
        return ExcelConstants.COVER_EN.equals(sheetName) || ExcelConstants.COVER_CN.equals(sheetName);
    }

    public static String getBaseName(File file) {
        String fileName = file.getName();
        int lastDotIndex = fileName.lastIndexOf(46);
        return lastDotIndex == -1 ? fileName : fileName.substring(0, lastDotIndex);
    }
}