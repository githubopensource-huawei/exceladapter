/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.main;

import com.app.excelcsvconverter.converter.CsvToExcelConverter;
import com.app.excelcsvconverter.converter.CustomPackage;
import com.app.excelcsvconverter.converter.ExcelToCsvConverter;
import com.app.excelcsvconverter.converter.MacroInjection;
import com.app.excelcsvconverter.genexcelstyle.ExcelStyleGenerator;
import com.app.excelcsvconverter.resultmodel.ResultData;
import com.app.excelcsvconverter.util.MessageUtil;
import com.google.common.base.Strings;

import java.io.File;

/**
 * 功能描述：
 *
 * @since 2019-08-07
 */
public class Main {

    public static String language = "EN";

    /**
     * 配置Excel转CSV功能支持的最大阈值，超过该值，将丢失样式文件
     */
    public static long DROP_STYLE_TYPE_THRESHOLD = 50 * 1024 * 1024;

    private static boolean isEnhance = false;

    private static SwingComponent swingComponent = null;

    public static SwingComponent getSwingComponent() {
        return swingComponent;
    }

    public static void main(String[] args) {
        if (args.length == 3) {
            language = args[0];
            if ("enhance".equals(args[1])) {
                isEnhance = true;
            }
            Integer threshold = Integer.valueOf(args[2]);
            if (threshold < 0) {
                threshold = 50;
            }
            DROP_STYLE_TYPE_THRESHOLD = threshold * 1024 * 1024;
        }
        swingComponent = new SwingComponent(language);
        swingComponent.init(isEnhance);
    }

    public ResultData operation(String tmpPath, int type) {
        if (type == SwingComponent.CSV_TO_EXECL) {
            CsvToExcelConverter csvToExcelConverter = new CsvToExcelConverter(language);
            if (!tmpPath.endsWith(".zip")) {
                return new ResultData(false,
                    MessageUtil.getMessage("FILE_FORMAT_INCORRECT_EN", "FILE_FORMAT_INCORRECT_CN", language));
            }
            return csvToExcelConverter.converter(tmpPath);
        } else if (type == SwingComponent.EXECL_TO_CSV) {
            ExcelToCsvConverter excelToCsvConverter = new ExcelToCsvConverter(language);
            if (!tmpPath.endsWith(".csv") && !tmpPath.endsWith(".xlsx") && !tmpPath.endsWith(".xlsm") && !new File(
                tmpPath).isDirectory()) {
                return new ResultData(false,
                    MessageUtil.getMessage("FILE_FORMAT_INCORRECT_EN", "FILE_FORMAT_INCORRECT_CN", language));
            }
            return excelToCsvConverter.converterExcel2Csv(tmpPath);
        } else if (type == SwingComponent.INJECT_MECRO_TO_EXCEL) {
            MacroInjection macroInjection = new MacroInjection();
            if (!tmpPath.endsWith(".xlsx")) {
                return new ResultData(false,
                    MessageUtil.getMessage("FILE_FORMAT_INCORRECT_EN", "FILE_FORMAT_INCORRECT_CN", language));
            }
            return macroInjection.injectMacro(tmpPath, language);
        } else if (type == SwingComponent.CUSTOM_PACKAGE_UPDATE) {
            CustomPackage customPackage = new CustomPackage();
            if (!tmpPath.endsWith(".zip") && !tmpPath.endsWith(".tar")) {
                return new ResultData(false,
                    MessageUtil.getMessage("FILE_FORMAT_INCORRECT_EN", "FILE_FORMAT_INCORRECT_CN", language));
            }
            return customPackage.updateCustomPackage(tmpPath, language);
        }
        return new ResultData(false,
            MessageUtil.getMessage("OPERATING_MODE_NOT_EXIST_EN", "OPERATING_MODE_NOT_EXIST_CN", language));
    }

    public ResultData operation(String srcPath, String destPath, int type) {
        if (type == SwingComponent.EXCEL_TO_STYLE) {
            if (!srcPath.endsWith(".xls") && !srcPath.endsWith(".xlsx") && !srcPath.endsWith(".xlsm")) {
                return new ResultData(false,
                    MessageUtil.getMessage("FILE_FORMAT_INCORRECT_EN", "FILE_FORMAT_INCORRECT_CN", language));
            }

            if (Strings.isNullOrEmpty(destPath)) {
                return new ResultData(false, "Please select export path.");
            }

            ExcelStyleGenerator excelStyleGenerator = new ExcelStyleGenerator();
            return excelStyleGenerator.generate(srcPath, destPath);
        }
        return new ResultData(false,
            MessageUtil.getMessage("OPERATING_MODE_NOT_EXIST_EN", "OPERATING_MODE_NOT_EXIST_CN", language));
    }
}