/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.converter;

import static com.app.excelcsvconverter.consts.summary.SummaryBookConst.COVER_SHEET_NAME;
import static com.app.excelcsvconverter.consts.summary.SummaryBookConst.COVER_SHEET_NAME_CN;
import static com.app.excelcsvconverter.util.ConfigUtil.CSV_SUFFIX_NAME;
import static com.app.excelcsvconverter.util.ConfigUtil.FILEIDENTIFICATION_CSV_NAME;
import static com.app.excelcsvconverter.util.ConfigUtil.MACRO_TOOL_PATH;

import com.app.excelcsvconverter.csv.CsvSheet;
import com.app.excelcsvconverter.main.Main;
import com.app.excelcsvconverter.parser.ExcelFormat;
import com.app.excelcsvconverter.parser.Xls2Csv;
import com.app.excelcsvconverter.parser.Xlsx2Csv;
import com.app.excelcsvconverter.resultmodel.ResultData;
import com.app.excelcsvconverter.util.ConfigUtil;
import com.app.excelcsvconverter.util.EncodingUtil;
import com.app.excelcsvconverter.util.ExcelUtil;
import com.app.excelcsvconverter.util.FileUtil;
import com.app.excelcsvconverter.util.MessageUtil;
import com.google.common.io.Files;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;

/**
 * 功能描述：
 *
 * @since 2019-08-18
 */
public class MacroInjection {

    private static final Logger LOGGER = LoggerFactory.getLogger(MacroInjection.class);

    private static final int TIME_OUT_COUNT_LIMIT = 90;

    private static final int TIME_OUT_TIME_LIMIT = 3000;

    private static final String LOG_FILE_END_FLAG = "end inject macro";

    private static final String CSV_SUFFIX = ".csv";

    private static final String INJECT_SUCCESS_FLAG = "inject macro successful";

    private static final String INJECT_FAILD_FLAG = "inject macro failed";

    private static final String INJECT_LOG_PATH = "injectmacro" + File.separator + "injectmacro.log";

    private static final String INJECT_PARAM_CSV_NAME = "injectParam";

    private static final String INJECT_COMMEND_NAME = "powershell.exe";

    private static final String INJECT_POWERSHELL_NAME = "injectmacro" + File.separator + "inject-excelmacro.ps1";

    private static final String INJECT_PARAM_PATH = "injectmacro" + File.separator + INJECT_PARAM_CSV_NAME + ".txt";

    private static final int INJECT_FAILD_TIMEOUT = 1;

    private static final int INJECT_SUCCESS = 2;

    private static final int INJECT_FAILD_FILEERROR = 3;

    private String tableType = "";

    public ResultData injectMacro(String excelPath) {
        return injectMacro(excelPath, Main.language);
    }

    public ResultData injectMacro(String excelPath, String language) {
        List<String> cmds = new ArrayList<>(4);
        String command = INJECT_COMMEND_NAME;
        File commandScript = new File(INJECT_POWERSHELL_NAME);
        if (!commandScript.exists() || !commandScript.isFile()) {
            return new ResultData(false,
                MessageUtil.getMessage("INJECT_MACRO_FAILED_EN", "INJECT_MACRO_FAILED_CN", language));
        }
        String commandScriptPath = "";
        try {
            commandScriptPath = commandScript.getCanonicalPath();
        } catch (IOException e) {
            return new ResultData(false,
                MessageUtil.getMessage("INJECT_MACRO_FAILED_EN", "INJECT_MACRO_FAILED_CN", language));
        }

        File excelFile = new File(excelPath);
        if (!excelFile.exists() || !excelFile.isFile()) {
            return new ResultData(false, MessageUtil.getMessage("FILE_NOT_EXIST_EN", "FILE_NOT_EXIST_CN", language));
        }

        List<String> pathList = getMacroPath(excelFile);
        if (pathList == null) {
            return new ResultData(false,
                MessageUtil.getMessage("MACRO_FILE_NOT_EXIST_EN", "MACRO_FILE_NOT_EXIST_CN", language));
        }
        File macroFile = new File(pathList.get(0));
        if (!macroFile.exists()) {
            return new ResultData(false,
                MessageUtil.getMessage("MACRO_FILE_NOT_EXIST_EN", "MACRO_FILE_NOT_EXIST_CN", language));
        }
        File paramFile = new File(INJECT_PARAM_PATH);
        if (!paramFile.exists()) {
            try {
                if (!paramFile.createNewFile()) {
                    LOGGER.error("create file failed");
                    return new ResultData(false,
                        MessageUtil.getMessage("INJECT_MACRO_FAILED_EN", "INJECT_MACRO_FAILED_EN", language));
                }
            } catch (IOException e) {
                return new ResultData(false,
                    MessageUtil.getMessage("INJECT_MACRO_FAILED_EN", "INJECT_MACRO_FAILED", language));
            }
        }

        String macroPath = "";
        try {
            macroPath = macroFile.getCanonicalPath();
        } catch (IOException e) {
            return new ResultData(false,
                MessageUtil.getMessage("INJECT_MACRO_FAILED_EN", "INJECT_MACRO_FAILED_CN", language));
        }
        File tempDir = null;
        String templatePath = "";
        File templateFile = new File(pathList.get(1));
        File tmpTemplateFile;
        if (templateFile.exists()) {
            try {
                tempDir = Files.createTempDir();
                FileUtil.copyFileToDirectory(templateFile, tempDir);
                tmpTemplateFile = new File(tempDir, templateFile.getName());
                if (tmpTemplateFile != null && tmpTemplateFile.exists()) {
                    templatePath = tmpTemplateFile.getCanonicalPath();
                }
            } catch (IOException e) {
                return new ResultData(false,
                    MessageUtil.getMessage("INJECT_MACRO_FAILED_EN", "INJECT_MACRO_FAILED_CN", language));
            }
        }
        writeParam(excelPath, macroPath, templatePath, paramFile);
        cmds.add(command);
        cmds.add("-ExecutionPolicy");
        cmds.add("UnRestricted");
        cmds.add("-File");
        cmds.add(commandScriptPath);

        //清空日志文件
        File logFile = new File(INJECT_LOG_PATH);
        if (logFile.exists()) {
            FileUtil.deleteFileQuietly(logFile);
        }
        try {
            new ProcessBuilder(cmds).start();
        } catch (IOException e) {
            e.printStackTrace();
            FileUtil.deleteDirectoryQuietly(tempDir);
            return new ResultData(false,
                MessageUtil.getMessage("INJECT_MACRO_FAILED_EN", "INJECT_MACRO_FAILED_CN", language));
        }

        int injectResult = getInjectResult(logFile);
        FileUtil.deleteDirectoryQuietly(tempDir);
        if (injectResult == INJECT_SUCCESS) {
            return new ResultData(true,
                MessageUtil.getMessage("INJECT_MACRO_SUCCESS_EN", "INJECT_MACRO_SUCCESS_CN", language));
        } else if (injectResult == INJECT_FAILD_TIMEOUT) {
            return new ResultData(true,
                MessageUtil.getMessage("INJECT_MACRO_FAILED_TIMEOUT_EN", "INJECT_MACRO_FAILED_TIMEOUT_CN", language));
        } else if (injectResult == INJECT_FAILD_FILEERROR) {
            return new ResultData(true,
                MessageUtil.getMessage("INJECT_MACRO_FAILED_FILEERROR_EN", "INJECT_MACRO_FAILED_FILEERROR_CN",
                    language));
        }
        return new ResultData(false,
            MessageUtil.getMessage("INJECT_MACRO_FAILED_EN", "INJECT_MACRO_FAILED_CN", language));
    }

    private void writeParam(String destPath, String macroPath, String templatePath, File filePath) {
        BufferedWriter out = null;
        try {
            out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(filePath, false), "UTF-8"));
            out.write("params" + "***");
            out.write(destPath + "***");
            out.write(macroPath + "***");
            out.write(templatePath + "***");
            out.write(tableType);
            out.flush();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private int getInjectResult(File file) {
        boolean logFileEnd = isLogFileEnd(file);
        if (!logFileEnd) {
            return INJECT_FAILD_TIMEOUT;
        }
        if (isFileContainValue(INJECT_SUCCESS_FLAG, file)) {
            return INJECT_SUCCESS;
        }
        if (isFileContainValue(INJECT_FAILD_FLAG, file)) {
            return INJECT_FAILD_FILEERROR;
        }
        return -1;
    }

    private boolean isLogFileEnd(File file) {
        int count = 0;
        while (true) {
            if (isFileContainValue(LOG_FILE_END_FLAG, file)) {
                return true;
            } else {
                try {
                    count++;
                    if (count > TIME_OUT_COUNT_LIMIT) {
                        return false;
                    }
                    Thread.sleep(TIME_OUT_TIME_LIMIT);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
            }

        }
    }

    private boolean isFileContainValue(String value, File file) {
        if (!file.exists()) {
            return false;
        }
        try (FileReader fileReader = new FileReader(file)) {
            BufferedReader reader = new BufferedReader(fileReader);
            String line;

            while ((line = reader.readLine()) != null) {
                if (line.contains(value)) {
                    return true;
                }
            }
        } catch (IOException e) {
            return false;
        }
        return false;
    }

    private File parseCsv2TmpDir(File excelFile) {
        OPCPackage p = null;
        // 创建临时文件目录
        File tempDir = Files.createTempDir();
        try {
            ZipSecureFile.setMinInflateRatio(-1.0d);
            ExcelFormat format = ExcelUtil.checkFormat(excelFile);
            if (ExcelFormat.OLE2.equals(format)) {
                Xls2Csv process = new Xls2Csv(excelFile, tempDir, EncodingUtil.getDefaultEncoding(), new ArrayList<>(),
                    true, true, true);
                process.process();
            } else if (ExcelFormat.OOXML.equals(format)) {
                p = OPCPackage.open(excelFile.getPath(), PackageAccess.READ);
                Xlsx2Csv process = new Xlsx2Csv(p, tempDir, EncodingUtil.getDefaultEncoding(), new ArrayList<>(), true,
                    true, true);
                process.process();
            } else if (ExcelFormat.NA.equals(format)) {
                FileUtil.deleteDirectoryQuietly(tempDir);
                return null;
            }
        } catch (Exception e) {
            FileUtil.deleteDirectoryQuietly(tempDir);
            return null;
        } finally {
            if (null != p) {
                try {
                    p.close();
                } catch (IOException e) {
                    FileUtil.deleteDirectoryQuietly(tempDir);
                    return null;
                }
            }
        }
        return tempDir;
    }

    private CsvSheet getCoverCsvSheet(File tmpdir) {
        File[] files = tmpdir.listFiles();
        if (files == null || files.length == 0) {
            LOGGER.error("Gets csv file list is empty.");
            return null;
        }

        for (File file : files) {
            if (file.getName().equals(COVER_SHEET_NAME + CSV_SUFFIX) || file.getName()
                .equals(COVER_SHEET_NAME_CN + CSV_SUFFIX)) {
                return new CsvSheet(file.getName().substring(0, file.getName().length() - 4), file.getParentFile());
            }
        }
        return null;
    }

    private CsvSheet getFileIdentificationCsvSheet(File tmpdir) {
        File inTernalPath = null;
        try {
            inTernalPath = new File(tmpdir.getCanonicalPath() + File.separator + "Internal");
        } catch (IOException e) {
            FileUtil.deleteDirectoryQuietly(tmpdir);
            e.printStackTrace();
        }
        if (inTernalPath == null || !inTernalPath.exists()) {
            FileUtil.deleteDirectoryQuietly(tmpdir);
            return null;
        }
        File[] files = inTernalPath.listFiles();
        if (files == null || files.length == 0) {
            LOGGER.error("Gets csv file list is empty.");
            return null;
        }
        for (File file : files) {
            if (file.getName().equals(FILEIDENTIFICATION_CSV_NAME + CSV_SUFFIX_NAME)) {
                return new CsvSheet(FILEIDENTIFICATION_CSV_NAME, file.getParentFile());
            }
        }
        return null;
    }

    private List<String> getMacroPath(File excelFile) {
        File tmpdir = parseCsv2TmpDir(excelFile);
        if (tmpdir == null) {
            LOGGER.error("create tmp dir failed.");
            return null;
        }
        CsvSheet fileIdCsvSheet = getFileIdentificationCsvSheet(tmpdir);
        if (fileIdCsvSheet == null) {
            LOGGER.error("get  fileIdCsvSheet failed.");
            FileUtil.deleteDirectoryQuietly(tmpdir);
            return null;
        }
        List<String> list = ConfigUtil.getInfoFromFileIdCsv(fileIdCsvSheet);
        if (list == null) {
            LOGGER.error("get fileIdCsvSheet info failed.");
            FileUtil.deleteDirectoryQuietly(tmpdir);
            return null;
        }
        String title = list.get(0);
        tableType = list.get(1);

        Map<String, List<String>> fileCodeConfig = ConfigUtil.getFileCodeConfig();
        if (title == "" || fileCodeConfig.isEmpty()) {
            LOGGER.error("read FileCodeMapping.csv failed.");
            FileUtil.deleteDirectoryQuietly(tmpdir);
            return null;
        }

        List<String> pathList = new ArrayList<>();
        String finalTitle = title;
        Optional<Map.Entry<String, List<String>>> optional = fileCodeConfig.entrySet()
            .stream()
            .filter(e -> e.getKey().equals(finalTitle))
            .findAny();
        if (optional.isPresent()) {
            Map.Entry<String, List<String>> pathSet = optional.get();
            List<String> value = pathSet.getValue();
            if (value.get(0) != null && !value.get(0).isEmpty()) {
                pathList.add(MACRO_TOOL_PATH + value.get(0));
            } else {
                pathList.add("");
            }
            if (value.get(1) != null && !value.get(1).isEmpty()) {
                CsvSheet coverCsvSheet = getCoverCsvSheet(tmpdir);
                String finalTemplatePath = value.get(1);
                if (coverCsvSheet != null) {
                    finalTemplatePath = ConfigUtil.getFinalTemplatePath(finalTemplatePath, tableType, coverCsvSheet);
                }

                pathList.add(MACRO_TOOL_PATH + finalTemplatePath);
            } else {
                pathList.add("");
            }
        } else {
            FileUtil.deleteDirectoryQuietly(tmpdir);
            return null;
        }
        FileUtil.deleteDirectoryQuietly(tmpdir);
        return pathList;
    }
}
