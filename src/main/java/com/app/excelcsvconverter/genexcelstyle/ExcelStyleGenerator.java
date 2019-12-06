/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.genexcelstyle;

import com.app.excelcsvconverter.resultmodel.ResultData;
import com.app.excelcsvconverter.writer.JaxbWriter;
import com.app.excelcsvconverter.xmlmodel.styledata.ExcelStyle;
import com.google.common.base.Strings;

import java.io.File;
import java.io.IOException;

/**
 * Excel样式文件生成
 *
 * @author s00202098
 * @since 2019-08-15
 */
public class ExcelStyleGenerator {

    public ExcelStyleGenerator() {

    }

    public ResultData generate(String excelFilePath, String destPath) {
        try {
            File resultFile = new File(destPath + File.separator + "style.xml");
            if (resultFile.exists()) {
                if (!resultFile.delete()) {
                    throw new IOException();
                }
            }
            if (!resultFile.createNewFile()) {
                throw new IOException();
            }

            ExcelReaderWithPOI excelReaderWithPOI = new ExcelReaderWithPOI();
            ExcelStyle excelStyle = excelReaderWithPOI.getExcelStyle(new File(excelFilePath));

            JaxbWriter.writeConfig(excelStyle, resultFile);
        } catch (IOException e) {
            e.printStackTrace();
            if (Strings.isNullOrEmpty(e.getMessage())) {
                return new ResultData(false, "Operation failed, maybe some exception happened.");
            }
            return new ResultData(false, e.getMessage());
        } catch (IllegalArgumentException e) {
            if (Strings.isNullOrEmpty(e.getMessage())) {
                return new ResultData(false, "Operation failed, maybe some exception happened.");
            }
            return new ResultData(false, e.getMessage());
        } catch (Exception e) {
            if (Strings.isNullOrEmpty(e.getMessage())) {
                return new ResultData(false, "Operation failed, maybe some exception happened.");
            }
            return new ResultData(false, e.getMessage());
        }

        return new ResultData(true, "Operation succeeded.");
    }
}
