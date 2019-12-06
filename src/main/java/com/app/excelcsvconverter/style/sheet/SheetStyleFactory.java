/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.style.sheet;

import com.app.excelcsvconverter.xmlmodel.styledata.SheetStyle;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

/**
 * 创建页签级样式的工厂
 *
 * @since 2019-07-22
 */
public class SheetStyleFactory {

    private List<SheetStyleOperator> sheetStyleOperators = new ArrayList<>();

    private List<SheetStyleOperator> getSheetStyleOperators() {
        ValidationSetter validationSetter = new ValidationSetter();
        sheetStyleOperators.add(validationSetter);
        FreezePaneSetter freezePaneSetter = new FreezePaneSetter();
        sheetStyleOperators.add(freezePaneSetter);
        MergeRegionSetter mergeRegionSetter = new MergeRegionSetter();
        sheetStyleOperators.add(mergeRegionSetter);
        TabColorSetter tabColorSetter = new TabColorSetter();
        sheetStyleOperators.add(tabColorSetter);
        DisplayGridlinesSetter displayGridlinesSetter = new DisplayGridlinesSetter();
        sheetStyleOperators.add(displayGridlinesSetter);
        return sheetStyleOperators;
    }

    public void setSheetLevelStyle(Sheet sheet, SheetStyle sheetStyle) {
        if (sheetStyle == null) {
            return;
        }
        for (SheetStyleOperator sheetStyleOperator : getSheetStyleOperators()) {
            sheetStyleOperator.setStyle(sheet, sheetStyle);
        }
    }
}