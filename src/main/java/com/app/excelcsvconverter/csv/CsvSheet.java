/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.csv;

import com.app.excelcsvconverter.parser.CloseableIterable;
import com.app.excelcsvconverter.util.FileEncodingDetector;
import com.google.common.io.Files;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 存放sheet相关信息的实体
 *
 * @since 2019-08-08
 */
public class CsvSheet {
    private static final Logger LOGGER = LoggerFactory.getLogger(CsvSheet.class);

    private static final String CSV_SUFFIX = ".csv";

    private String sheetName;

    private File tempCsvPath;

    public CsvSheet(String sheetName, File tempCsvPath) {
        this.sheetName = sheetName;
        this.tempCsvPath = tempCsvPath;
    }

    public String getSheetName() {
        return sheetName;
    }

    public CloseableIterable iterator() throws IOException {
        File csvFile = new File(tempCsvPath, sheetName + CSV_SUFFIX);
        BufferedReader fileReader = Files.newReader(csvFile, Charset.forName(FileEncodingDetector.detect(csvFile)));
        CSVFormat format = CSVFormat.DEFAULT.withIgnoreEmptyLines(false);
        CSVParser parser = format.parse(fileReader);
        Iterator<CSVRecord> iterator = parser.iterator();
        return new CloseableIterable() {

            @Override
            public boolean hasNext() {
                return iterator.hasNext();
            }

            @Override
            public List<String> next() {
                CSVRecord csvRecord = iterator.next();
                List<String> valueList = new ArrayList<>();
                Iterator<String> iteratorStr = csvRecord.iterator();
                while(iteratorStr.hasNext()) {
                    valueList.add(iteratorStr.next());
                }
                return valueList;
            }

            @Override
            public void close() {
                if (fileReader != null) {
                    try {
                        fileReader.close();
                    } catch (Exception e) {
                        LOGGER.error("close fileReader failed.", e);
                    }
                }
            }
        };
    }
}