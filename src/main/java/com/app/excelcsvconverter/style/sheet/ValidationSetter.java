/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.sheet;

import com.app.excelcsvconverter.util.ExcelConstants;
import com.app.excelcsvconverter.xmlmodel.styledata.CellRange;
import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;
import com.app.excelcsvconverter.xmlmodel.styledata.Validation;
import com.google.common.base.Strings;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.List;

/**
 * 设置下拉框
 *
 * @since 2019-07-22
 */
public class ValidationSetter implements SheetStyleOperator {

    private static final String COLON = ":";

    @Override
    public void setStyle(Sheet sheet, SheetStyle sheetStyle) {
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        if (dvHelper == null) {
            return;
        }
        List<Validation> validations = sheetStyle.getValidation();
        for (Validation validation : validations) {
            CellRange cellRange = validation.getCellRange();
            if (cellRange == null) {
                continue;
            }
            if (Strings.isNullOrEmpty(validation.getAttrs())) {
                continue;
            }

            formatFormula(validation);

            DataValidationConstraint dvConstraint;
            if (validation.getAttrs2() == null) {
                dvConstraint = validation.getAttrs().startsWith("=")
                    ? dvHelper.createFormulaListConstraint(validation.getAttrs())
                    : dvHelper.createExplicitListConstraint(validation.getAttrs().split(ExcelConstants.SPLIT_REGEX));
            } else {
                dvConstraint = createDataValidationConstraint(validation, dvHelper);
            }
            CellRangeAddressList addressList = buildCellRangeAddressList(cellRange);
            DataValidation dataValidation = dvHelper.createValidation(dvConstraint, addressList);
            dataValidation.setSuppressDropDownArrow(true);
            String errorBoxTitle = validation.getErrorBoxTitle();
            String errorBoxText = validation.getErrorBoxText();
            if (!Strings.isNullOrEmpty(errorBoxText)) {
                dataValidation.createErrorBox(errorBoxTitle, errorBoxText);
                dataValidation.setShowErrorBox(true);
            } else {
                dataValidation.setShowErrorBox(false);
            }
            String promptBoxText = validation.getPromptBoxText();
            if (!Strings.isNullOrEmpty(promptBoxText)) {
                dataValidation.createPromptBox(validation.getPromptBoxTitle(), promptBoxText);
                dataValidation.setShowPromptBox(true);
            } else {
                dataValidation.setShowPromptBox(false);
            }
            sheet.addValidationData(dataValidation);
        }
    }

    private void formatFormula(Validation validation) {
        if (validation.getValidationType() == DataValidationConstraint.ValidationType.LIST) {
            String listStr = validation.getAttrs();
            if (listStr.startsWith("=")) {
                return;
            }
            if (listStr.contains(COLON)) {
                validation.setAttrs("=" + listStr);
            }
        }
    }

    /**
     * 构建CellRangeAddressList对象
     *
     * @return
     */
    private CellRangeAddressList buildCellRangeAddressList(CellRange cellRange) {
        return new CellRangeAddressList(cellRange.getRowBegin(), cellRange.getRowEnd(), cellRange.getColBegin(),
            cellRange.getColEnd());
    }

    private DataValidationConstraint createDataValidationConstraint(Validation validation,
        DataValidationHelper dvHelper) {
        int validationType = validation.getValidationType();
        switch (validationType) {
            case DataValidationConstraint.ValidationType.INTEGER:
                return dvHelper.createIntegerConstraint(validation.getOperator(), validation.getAttrs(),
                    validation.getAttrs2());
            case DataValidationConstraint.ValidationType.DECIMAL:
                return dvHelper.createDecimalConstraint(validation.getOperator(), validation.getAttrs(),
                    validation.getAttrs2());
            case DataValidationConstraint.ValidationType.LIST:
                return validation.getAttrs().startsWith("=")
                    ? dvHelper.createFormulaListConstraint(validation.getAttrs())
                    : dvHelper.createExplicitListConstraint(validation.getAttrs().split(ExcelConstants.SPLIT_REGEX));
            case DataValidationConstraint.ValidationType.DATE:
                return dvHelper.createDateConstraint(validation.getOperator(), validation.getAttrs(),
                    validation.getAttrs2(), ExcelConstants.YYYY_MM_DD);
            case DataValidationConstraint.ValidationType.TIME:
                return dvHelper.createTimeConstraint(validation.getOperator(), validation.getAttrs(),
                    validation.getAttrs2());
            case DataValidationConstraint.ValidationType.TEXT_LENGTH:
                return dvHelper.createTextLengthConstraint(validation.getOperator(), validation.getAttrs(),
                    validation.getAttrs2());
            default:
                return dvHelper.createCustomConstraint(validation.getAttrs());
        }
    }
}