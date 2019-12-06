/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.converter;

import com.app.excelcsvconverter.csv.CsvSheet;
import com.app.excelcsvconverter.genexcelstyle.ExcelStyleGenerator;
import com.app.excelcsvconverter.main.Main;
import com.app.excelcsvconverter.parser.CloseableIterable;
import com.app.excelcsvconverter.parser.ExcelFormat;
import com.app.excelcsvconverter.parser.Xls2Csv;
import com.app.excelcsvconverter.parser.Xlsx2Csv;
import com.app.excelcsvconverter.resultmodel.ResultData;
import com.app.excelcsvconverter.util.CompressUtil;
import com.app.excelcsvconverter.util.ExcelUtil;
import com.app.excelcsvconverter.util.FileUtil;
import com.app.excelcsvconverter.util.MessageUtil;
import com.google.common.io.Files;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel To Csv转换器
 *
 * @since 2019-07-31
 */
public class ExcelToCsvConverter {

    private final Logger LOGGER = LoggerFactory.getLogger(ExcelToCsvConverter.class);

    private final String ZIP_SUFFIX_NAME = ".zip";

    private final String SHEET_STYLE = "SHEET STYLE.csv";

    private String language;

    private boolean operateDirSuccess = true;

    private String operateDirErrorFile = "";

    private boolean isInputMultiFile = false;

    public ExcelToCsvConverter(String language) {
        this.language = language;
    }

    public ResultData converterExcel2Csv(String excelFilePath) {
        File file = new File(excelFilePath);
        if (file.isDirectory()) {
            isInputMultiFile = true;
            if (!file.exists()) {
                return new ResultData(false,
                    MessageUtil.getMessage("EXCEL2CSV_DIR_FAILED_EN", "EXCEL2CSV_DIR_FAILED_CN", language)
                        + excelFilePath);
            }
            converterAdapter(excelFilePath);
            if (operateDirSuccess) {
                return new ResultData(true,
                    MessageUtil.getMessage("EXCEL_TO_CSV_SUCCESS_EN", "EXCEL_TO_CSV_SUCCESS_CN", language));
            } else {
                return new ResultData(false,
                    MessageUtil.getMessage("EXCEL2CSV_FILE_FAILED_EN", "EXCEL2CSV_FILE_FAILED_CN", language)
                        + operateDirErrorFile);
            }
        } else if (excelFilePath.endsWith(".csv")) {
            isInputMultiFile = true;
            File csvFile = new File(excelFilePath);
            List<String> fileList = new ArrayList<>();
            CsvSheet csvSheet = new CsvSheet(csvFile.getName().substring(0, csvFile.getName().length() - 4),
                csvFile.getParentFile());

            try (CloseableIterable iterator = csvSheet.iterator()) {
                while (iterator.hasNext()) {
                    List<String> next = iterator.next();
                    fileList.add(next.get(0));
                }
            } catch (Exception e) {
                LOGGER.error("save: Obtain iterator failed.", e);
            }
            for (String path : fileList) {
                ResultData resultData = converterExcel2Csv(path);
                if (!resultData.isSuccessful()) {
                    return resultData;
                }
            }

            return new ResultData(true,
                MessageUtil.getMessage("EXCEL_TO_CSV_SUCCESS_EN", "EXCEL_TO_CSV_SUCCESS_CN", language));
        } else {
            return converter(excelFilePath);
        }
    }

    private void converterAdapter(String excelFilePath) {
        if (!operateDirSuccess) {
            return;
        }
        File file = new File(excelFilePath);
        if (!file.isDirectory()) {
            if (!excelFilePath.endsWith(".xlsx") && !excelFilePath.endsWith(".xlsm")) {
                return;
            }
            ResultData converter = converter(excelFilePath);
            if (!converter.isSuccessful()) {
                operateDirSuccess = false;
                operateDirErrorFile = excelFilePath;

            }
        } else {
            File[] files = file.listFiles();
            if (files == null || files.length == 0) {
                return;
            }
            for (File f : files) {
                try {
                    converterAdapter(f.getCanonicalPath());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private void copySheetStyleToInternalDir(File outDir) {
        File internalFile = new File(outDir, "Internal");
        if (!internalFile.exists()) {
            outDir.mkdirs();
        }
        File sheetStyleFile = new File(outDir, SHEET_STYLE);
        if (sheetStyleFile.exists()) {
            try {
                FileUtil.copyFileToDirectory(sheetStyleFile, internalFile);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                FileUtil.deleteFileQuietly(sheetStyleFile);
            }
        }
    }

    private ResultData converter(String excelFilePath) {
        File file = new File(excelFilePath);
        if (!file.exists() || !file.isFile()) {
            return new ResultData(false, MessageUtil.getMessage("FILE_NOT_EXIST_EN", "FILE_NOT_EXIST_CN", language));
        }
        long fileLen = file.length();
        if (fileLen > Main.DROP_STYLE_TYPE_THRESHOLD) {
            boolean result = true;
            if (!isInputMultiFile) {
                result = Main.getSwingComponent()
                    .confirmDialog(MessageUtil.getMessage("OPTION_DIALOG_TITLE_EN", "OPTION_DIALOG_TITLE_CN", language),
                        MessageUtil.getMessage("OPTION_DIALOG_MESSAGE_EN", "OPTION_DIALOG_MESSAGE_CN", language));
            }

            if (!result) {
                return null;
            }
        }
        File tempDir = Files.createTempDir();
        File outDir = new File(tempDir, file.getName().substring(0, file.getName().indexOf(".")));
        if (!outDir.exists()) {
            outDir.mkdirs();
        }
        OPCPackage p = null;
        try {
            ZipSecureFile.setMinInflateRatio(-1.0d);
            ExcelFormat format = ExcelUtil.checkFormat(file);
            if (ExcelFormat.OLE2.equals(format)) {
                Xls2Csv process = new Xls2Csv(file, outDir, StandardCharsets.UTF_8, new ArrayList<>(), true, true,
                    false);
                process.process();
            } else if (ExcelFormat.OOXML.equals(format)) {
                p = OPCPackage.open(file.getPath(), PackageAccess.READ);
                Xlsx2Csv process = new Xlsx2Csv(p, outDir, StandardCharsets.UTF_8, new ArrayList<>(), true, true,
                    false);
                process.process();
            } else if (ExcelFormat.NA.equals(format)) {
                FileUtil.deleteDirectoryQuietly(tempDir);
                return new ResultData(false,
                    MessageUtil.getMessage("FILE_NOT_MATCH_NAME_EN", "FILE_NOT_MATCH_NAME_CN", language));
            }

        } catch (Exception var13) {
            FileUtil.deleteDirectoryQuietly(tempDir);
            LOGGER.error("Error transform to csv: ", var13);
            return new ResultData(false,
                MessageUtil.getMessage("EXCEL_TO_CSV_FAIL_EN", "EXCEL_TO_CSV_FAIL_CN", language));
        } finally {
            if (null != p) {
                try {
                    p.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        copySheetStyleToInternalDir(outDir);

        if (fileLen <= Main.DROP_STYLE_TYPE_THRESHOLD) {
            // 生成style.xml文件
            File styleFile = new File(outDir, "Style");
            if (styleFile.mkdir()) {
                ExcelStyleGenerator generator = new ExcelStyleGenerator();
                ResultData resultData = generator.generate(excelFilePath, styleFile.getPath());
                if (!resultData.isSuccessful()) {
                    return resultData;
                }
            }
        }

        File sourceFile = new File(outDir.getPath() + ZIP_SUFFIX_NAME);

        try {
            CompressUtil.zipFiles(outDir.getPath(), sourceFile, true);
        } catch (Exception e) {
            FileUtil.deleteDirectoryQuietly(tempDir);
            return new ResultData(false,
                MessageUtil.getMessage("EXCEL_TO_CSV_FAIL_EN", "EXCEL_TO_CSV_FAIL_CN", language));
        }

        try {
            String finalFilePath = file.getPath().substring(0, file.getPath().lastIndexOf(".")) + ZIP_SUFFIX_NAME;
            File finalFile = new File(finalFilePath);
            Files.copy(sourceFile, finalFile);
            FileUtil.deleteDirectoryQuietly(sourceFile.getParentFile());
            FileUtil.deleteDirectoryQuietly(tempDir);
        } catch (IOException e) {
            FileUtil.deleteDirectoryQuietly(tempDir);
            return new ResultData(false,
                MessageUtil.getMessage("EXCEL_TO_CSV_FAIL_EN", "EXCEL_TO_CSV_FAIL_CN", language));
        }
        return new ResultData(true,
            MessageUtil.getMessage("EXCEL_TO_CSV_SUCCESS_EN", "EXCEL_TO_CSV_SUCCESS_CN", language));
    }
}