package com.app.excelcsvconverter.util;

import com.app.excelcsvconverter.converter.TemplateType;
import com.app.excelcsvconverter.csv.CsvSheet;
import com.app.excelcsvconverter.parser.CloseableIterable;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 功能描述：
 *
 * @since 2019-09-18
 */
public class ConfigUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(ConfigUtil.class);

    private static Map<String, List<String>> configMap = new HashMap<>();

    public static final String MACRO_TOOL_PATH = "Data" + File.separator;

    public static final String FILEIDENTIFICATION_CSV_NAME = "FileIdentification";

    public static final String CSV_SUFFIX_NAME = ".csv";

    private static final String MACRO_TOOL_FILE = "FileCodeMapping";

    public static Map<String, List<String>> getFileCodeConfig() {
        if (configMap.isEmpty()) {
            File configPath = new File(MACRO_TOOL_PATH);
            CsvSheet csvSheet = new CsvSheet(MACRO_TOOL_FILE, configPath);

            try (CloseableIterable iterator = csvSheet.iterator()) {
                int rowIndex = 0;
                while (iterator.hasNext()) {
                    if (rowIndex == 0) {
                        iterator.next();
                        rowIndex++;
                        continue;
                    }

                    List<String> next = iterator.next();
                    if (next == null || next.size() < 2 || next.get(0) == null || next.get(0).isEmpty()) {
                        rowIndex++;
                        continue;
                    }
                    List<String> list = new ArrayList<>();
                    list.add(next.get(1) == null ? "" : next.get(1));
                    if (next.size() >= 3) {
                        list.add(next.get(2) == null ? "" : next.get(2));
                    } else {
                        list.add("");
                    }
                    configMap.put(next.get(0), list);

                    rowIndex++;
                }

            } catch (Exception e) {
                return new HashMap<>();
            }
        }
        return configMap;
    }

    public static List<String> getInfoFromFileIdCsv(CsvSheet fileIdCsvSheet) {
        List<String> result = new ArrayList<>();
        String title = "";
        String fileType = "";
        try (CloseableIterable iterator = fileIdCsvSheet.iterator()) {
            int rowIndex = 0;
            while (iterator.hasNext()) {
                List<String> next = iterator.next();
                if (rowIndex == 1) {
                    title = next.get(0);
                    if (next.size() > 1) {
                        fileType = next.get(1);
                    } else {
                        fileType = "";
                    }
                    break;
                }
                rowIndex++;
            }

        } catch (Exception e) {
            LOGGER.error("save: Obtain iterator failed.", e);
            return null;
        }
        if (title.isEmpty()) {
            return null;
        }
        result.add(title);
        result.add(fileType);
        return result;
    }

    public static String getFinalTemplatePath(String templatePath, String fileType, CsvSheet coverCsv) {
        if (fileType == null || fileType.trim().isEmpty() || !TemplateType.of(fileType).isControllerSummary()) {
            return templatePath;
        }
        String version = "";
        try (CloseableIterable iterator = coverCsv.iterator()) {
            int rowIndex = 0;
            while (iterator.hasNext()) {
                if (rowIndex == 2) {
                    List<String> next = iterator.next();
                    if (next.size() >= 3) {
                        version = next.get(3);
                    }
                    break;
                }
                rowIndex++;
            }

        } catch (Exception e) {
            LOGGER.error("save: Obtain iterator failed.", e);
            return null;
        }

        if (version != null || !version.isEmpty()) {
            return MessageFormat.format(templatePath, "CMEComponent" + "_" + version, version);
        }

        return templatePath;
    }
}
