/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.converter;

import com.app.excelcsvconverter.csv.CsvSheet;
import com.app.excelcsvconverter.parser.CloseableIterable;
import com.app.excelcsvconverter.resultmodel.ResultData;
import com.app.excelcsvconverter.util.CompressUtil;
import com.app.excelcsvconverter.util.FileUtil;
import com.app.excelcsvconverter.util.MessageUtil;
import com.google.common.io.Files;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 功能描述：
 *
 * @since 2019-08-19
 */
public class CustomPackage {
    private static final Logger LOGGER = LoggerFactory.getLogger(CustomPackage.class);

    /**
     * 临时文件夹下拷贝用户文件的临时文件夹目录
     */
    private final String CUSTOM_PACKAGE_FILE = "custom_package_file";

    /**
     * 临时文件夹下拷贝工具文件的临时文件夹目录
     */
    private final String TOOL_PACKAGE_FILE = "tool_package_file";

    /**
     * 工具中定制包路径
     */
    private final String TOOL_PACKAGE_CONFIG_PATH = "CustomizeTools";

    /**
     * 输入文件中字典信息
     */
    private final String DIC_FILE = "dict_R19SPC500.csv";

    private final String DIC_FILE_NAME = "dict_R19SPC500";

    private final String FILE_TYPE_ZIP = ".zip";

    private final String FILE_TYPE_TAR = ".tar";

    /**
     * zip包白名单
     */
    private static List<String> customZipWhiteList = new ArrayList<>();

    /**
     * tar包白名单
     */
    private static List<String> customTarWhiteList = new ArrayList<>();

    /**
     * 字典数据Map<原始的, 替换之后的>
     */
    private static Map<String, String> dicMap = new HashMap<>();

    public ResultData updateCustomPackage(String packagePath, String language) {

        File packageFile = new File(packagePath);
        if (!packageFile.exists()) {
            return new ResultData(false, MessageUtil.getMessage("FILE_NOT_EXIST_EN", "FILE_NOT_EXIST_CN", language));
        }

        // 创建临时文件目录
        File tempDir = Files.createTempDir();

        String fileName = packageFile.getName();
        ResultData resultData;
        dicMap = new HashMap<>();
        if (fileName.endsWith(FILE_TYPE_ZIP)) {
            try {
                resultData = updateZipFile(packageFile, language, tempDir);
            } catch (IOException e) {
                FileUtil.deleteDirectoryQuietly(tempDir);
                return new ResultData(false,
                    MessageUtil.getMessage("CUSTOM_PACKAGE_FAILED_EN", "CUSTOM_PACKAGE_FAILED_CN", language));
            }
            FileUtil.deleteDirectoryQuietly(tempDir);
            return resultData;
        } else if (fileName.endsWith(FILE_TYPE_TAR)) {
            try {
                resultData = updateTarFile(packageFile, language, tempDir);
            } catch (Exception e) {
                FileUtil.deleteDirectoryQuietly(tempDir);
                return new ResultData(false,
                    MessageUtil.getMessage("CUSTOM_PACKAGE_FAILED_EN", "CUSTOM_PACKAGE_FAILED_CN", language));
            }
            FileUtil.deleteDirectoryQuietly(tempDir);
            return resultData;
        }

        FileUtil.deleteDirectoryQuietly(tempDir);
        return new ResultData(true,
            MessageUtil.getMessage("CUSTOM_PACKAGE_SUCCESS_EN", "CUSTOM_PACKAGE_SUCCESS_CN", language));
    }

    /**
     * 查找用户包，对应的工具中的包目录
     *
     * @param fileName
     * @return
     */
    private String findToolPackagePath(String fileName) {
        fileName = changeFileName(fileName);
        return TOOL_PACKAGE_CONFIG_PATH + File.separator + fileName;
    }

    private ResultData updateZipFile(File packageFile, String language, File tempDir) throws IOException {
        File customTmpPath = new File(tempDir + File.separator + CUSTOM_PACKAGE_FILE);
        if (!customTmpPath.mkdir()) {
            return new ResultData(false,
                MessageUtil.getMessage("GENERATE_TEMPORARY_FILE_EN", "GENERATE_TEMPORARY_FILE_CN", language));
        }
        /**
         * 用户输入的zip文件解压之后在文件路径
         */
        String unzipCustomFilePath;
        try {
            unzipCustomFilePath = CompressUtil.unZipFiles(packageFile, customTmpPath.getCanonicalPath());
        } catch (Exception e) {
            return new ResultData(false,
                MessageUtil.getMessage("COMPRESS_FILE_INCORRECT_EN", "COMPRESS_FILE_INCORRECT_CN", language));
        }
        File unzipCustomFile = new File(unzipCustomFilePath);

        File[] unzipCustomFiles = unzipCustomFile.listFiles();

        // 去除标签文件
        for (File file : unzipCustomFiles) {
            if (file.getName().endsWith(".cms") || file.getName().endsWith(".crl")) {
                FileUtil.deleteFileQuietly(file);
            }
        }

        if (unzipCustomFiles == null || unzipCustomFiles.length == 0) {
            LOGGER.error("Gets csv file list is empty.");
            return new ResultData(false,
                MessageUtil.getMessage("UNZIP_DIRECTORY_EMPTY_EN", "UNZIP_DIRECTORY_EMPTY_CN", language));
        }
        int count = 0;
        for (File file : unzipCustomFiles) {
            if (isBelongZipWhiteList(file.getName())) {
                break;
            } else {
                count += 1;
            }
        }
        if (count >= unzipCustomFiles.length) {
            return new ResultData(false,
                MessageUtil.getMessage("CUSTOM_PACKAGE_INVALID_EN", "CUSTOM_PACKAGE_INVALID_CN", language));

        }

        initDicFile(unzipCustomFile);

        if (dicMap.isEmpty()) {
            return new ResultData(false, MessageUtil.getMessage("CONNOT_FIND_DIC_EN", "CONNOT_FIND_DIC_CN", language));
        }

        File toolExcelPackagePath = new File(findToolPackagePath(unzipCustomFiles[count].getName()));
        if (!toolExcelPackagePath.exists() || !toolExcelPackagePath.isDirectory()) {
            return new ResultData(false,
                MessageUtil.getMessage("CUSTOM_PACKAGE_EMPTY_EN", "CUSTOM_PACKAGE_EMPTY_CN", language));
        }

        File toolTmpPath = new File(tempDir + File.separator + TOOL_PACKAGE_FILE);
        if (!toolTmpPath.mkdir()) {
            return new ResultData(false,
                MessageUtil.getMessage("GENERATE_TEMPORARY_FILE_EN", "GENERATE_TEMPORARY_FILE_CN", language));
        }
        try {
            /**
             * 把工具目录中的文件复制到临时目录
             */
            FileUtil.copyDirectoryToDirectory(toolExcelPackagePath, toolTmpPath);
            recursiveTraversalFolder(toolTmpPath.getCanonicalPath());

        } catch (IOException e) {
            return new ResultData(false,
                MessageUtil.getMessage("CUSTOM_PACKAGE_EMPTY_EN", "CUSTOM_PACKAGE_EMPTY_CN", language));
        }

        // 合并定制包，相同文件以用户的包为准
        for (File file : unzipCustomFiles) {
            if (file.isFile()) {
                try {
                    FileUtil.copyFileToDirectory(file, toolTmpPath);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else if (file.isDirectory()) {
                try {
                    FileUtil.copyDirectoryToDirectory(file, toolTmpPath);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        try {
            CompressUtil.zipFiles(toolTmpPath.getCanonicalPath(), packageFile, true);
        } catch (Exception e) {
            return new ResultData(false,
                MessageUtil.getMessage("CUSTOM_PACKAGE_FAILED_EN", "CUSTOM_PACKAGE_FAILED_CN", language));
        }

        return new ResultData(true,
            MessageUtil.getMessage("CUSTOM_PACKAGE_SUCCESS_EN", "CUSTOM_PACKAGE_SUCCESS_CN", language));
    }

    /**
     * 加载字典数据
     *
     * @param file
     */
    private void initDicFile(File file) {
        if (file == null || !dicMap.isEmpty()) {
            return;
        }
        File[] files = file.listFiles();
        if (files == null || files.length == 0) {
            return;
        }
        int length = files.length;
        for (int index = 0; index < length; index++) {
            if (files[index].isDirectory()) {
                initDicFile(files[index]);
            } else {
                if (DIC_FILE.equals(files[index].getName())) {
                    CsvSheet csvSheet = new CsvSheet(DIC_FILE_NAME, files[index].getParentFile());
                    try (CloseableIterable iterator = csvSheet.iterator()) {
                        iterator.next();
                        while (iterator.hasNext()) {
                            List<String> next = iterator.next();
                            if (next.get(0) != null && next.get(0).isEmpty() || next.get(1) != null && next.get(1)
                                .isEmpty()) {
                                continue;
                            }
                            dicMap.put(next.get(0), next.get(1));
                        }
                    } catch (Exception e) {
                        LOGGER.error("save: Obtain iterator failed.", e);
                    }
                    /**
                     * 找到字典文件加载之后，可以删除字典文件
                     */
                    FileUtil.deleteFileQuietly(files[index]);
                }

            }
        }
    }

    /**
     * 将文件夹下的文件名称转换成非加密名称
     *
     * @param path
     */
    private void recursiveTraversalFolder(String path) {
        File folder = new File(path);
        if (!folder.exists()) {
            return;
        }

        File[] fileArr = folder.listFiles();
        if (null != fileArr && fileArr.length > 0) {
            for (File file : fileArr) {
                if (file.isDirectory()) {
                    File resFile = changeFile(file);
                    recursiveTraversalFolder(resFile.getAbsolutePath());
                } else {
                    changeFile(file);
                }
            }
        }

    }

    private File changeFile(File inputFile) {
        if (dicMap.isEmpty()) {
            return null;
        }
        String name = inputFile.getName();
        File parentFile = inputFile.getParentFile();
        for (Map.Entry<String, String> map : dicMap.entrySet()) {
            if (name.contains(map.getValue())) {
                String replace = name.replace(map.getValue(), map.getKey());
                File file = new File(parentFile + File.separator + replace);
                inputFile.renameTo(file);
                return file;
            }
        }
        return inputFile;
    }

    private String changeFileName(String fileName) {
        if (dicMap.isEmpty()) {
            return fileName;
        }

        for (Map.Entry<String, String> map : dicMap.entrySet()) {
            if (fileName.contains(map.getKey())) {
                return fileName.replace(map.getKey(), map.getValue());
            }
        }
        return fileName;
    }

    private ResultData updateTarFile(File packagePath, String language, File tempDir) throws Exception {
        File unTarTmpPath = CompressUtil.unTarFiles(packagePath, tempDir);
        int count = 0;
        File[] unTarTmpFiles = unTarTmpPath.listFiles();
        if (unTarTmpFiles == null || unTarTmpFiles.length == 0) {
            return new ResultData(false,
                MessageUtil.getMessage("CUSTOM_PACKAGE_INVALID_EN", "CUSTOM_PACKAGE_INVALID_CN", language));
        }
        // 去除标签文件
        for (File file : unTarTmpFiles) {
            if (file.getName().endsWith(".cms") || file.getName().endsWith(".crl")) {
                FileUtil.deleteFileQuietly(file);
            }
        }
        unTarTmpFiles = unTarTmpPath.listFiles();
        if (unTarTmpFiles == null || unTarTmpFiles.length == 0) {
            return new ResultData(false,
                MessageUtil.getMessage("CUSTOM_PACKAGE_INVALID_EN", "CUSTOM_PACKAGE_INVALID_CN", language));
        }
        String gzFileName = "";
        for (File file : unTarTmpFiles) {
            if (isBelongTarWhiteList(file.getName())) {
                gzFileName = file.getName();
                break;
            } else {
                count += 1;
            }
        }

        initDicFile(unTarTmpPath);

        if (count >= unTarTmpFiles.length) {
            return new ResultData(false,
                MessageUtil.getMessage("CUSTOM_PACKAGE_INVALID_EN", "CUSTOM_PACKAGE_INVALID_CN", language));

        }
        File inputFileTmpPath = new File(unTarTmpPath + File.separator + CUSTOM_PACKAGE_FILE);
        /**
         * tar包下还有tar.gz包，需要再解压
         */
        File unGzTmpPath = CompressUtil.unGzFiles(unTarTmpFiles[count], inputFileTmpPath);

        /**
         * 由于字典文件所在路径不固定，如果解压gz之前没有获取过，解压之后，需要再获取一次
         */
        if (dicMap.isEmpty()) {
            initDicFile(inputFileTmpPath);
        }

        if (dicMap.isEmpty()) {
            return new ResultData(false, MessageUtil.getMessage("CONNOT_FIND_DIC_EN", "CONNOT_FIND_DIC_CN", language));
        }

        File toolExcelPackagePath = new File(findToolPackagePath(unGzTmpPath.getName()));

        File toolGzTmpPath = new File(unTarTmpPath + File.separator + TOOL_PACKAGE_FILE);
        if (!toolGzTmpPath.mkdir()) {
            return new ResultData(false,
                MessageUtil.getMessage("GENERATE_TEMPORARY_FILE_EN", "GENERATE_TEMPORARY_FILE_CN", language));
        }
        try {
            FileUtil.copyDirectoryToDirectory(toolExcelPackagePath, toolGzTmpPath);
            recursiveTraversalFolder(toolGzTmpPath.getCanonicalPath());
        } catch (IOException e) {
            return new ResultData(false,
                MessageUtil.getMessage("CUSTOM_PACKAGE_EMPTY_EN", "CUSTOM_PACKAGE_EMPTY_CN", language));
        }

        // 合并定制包，相同文件以用户的包为准
        for (File file : unGzTmpPath.listFiles()) {
            if (file.isFile()) {
                try {
                    FileUtil.copyFileToDirectory(file, toolGzTmpPath.listFiles()[0]);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else if (file.isDirectory()) {
                try {
                    FileUtil.copyDirectoryToDirectory(file, toolGzTmpPath.listFiles()[0]);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        String createTarFileName = gzFileName.substring(0, gzFileName.length() - 3);
        CompressUtil.tarFiles(toolGzTmpPath.listFiles()[0],
            new File(unTarTmpPath + File.separator + createTarFileName));
        FileUtil.deleteDirectoryQuietly(new File(unTarTmpPath + File.separator + TOOL_PACKAGE_FILE));
        FileUtil.deleteDirectoryQuietly(new File(unTarTmpPath + File.separator + CUSTOM_PACKAGE_FILE));
        FileUtil.deleteFileQuietly(new File(unTarTmpPath + File.separator + gzFileName));
        CompressUtil.gzTar(new File(unTarTmpPath + File.separator + createTarFileName));
        CompressUtil.tarFiles(unTarTmpPath, packagePath);
        return new ResultData(true,
            MessageUtil.getMessage("CUSTOM_PACKAGE_SUCCESS_EN", "CUSTOM_PACKAGE_SUCCESS_CN", language));
    }

    private boolean isBelongZipWhiteList(String name) {
        if (customZipWhiteList.isEmpty()) {
            customZipWhiteList.add("CME_LLD_ConvertTool_FEATURE_CN");
            customZipWhiteList.add("CME_LLD_ConvertTool_FEATURE_EN");
        }
        return customZipWhiteList.contains(name);
    }

    private boolean isBelongTarWhiteList(String name) {
        if (customTarWhiteList.isEmpty()) {
            customTarWhiteList.add("CME_IUB_Convert_Tool_FEATURE.tar.gz");
            customTarWhiteList.add("CME_VDFD2_GERMANYTMO_FEATURE.tar.gz");
        }
        return customTarWhiteList.contains(name);
    }
}
