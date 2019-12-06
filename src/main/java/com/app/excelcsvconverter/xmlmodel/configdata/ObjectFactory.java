/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.xmlmodel.configdata;

import javax.xml.bind.annotation.XmlRegistry;

@XmlRegistry
public class ObjectFactory {
    public ObjectFactory() {
    }

    public CommonData createCommonData() {
        return new CommonData();
    }

    public ConfigData createConfigData() {
        return new ConfigData();
    }
}