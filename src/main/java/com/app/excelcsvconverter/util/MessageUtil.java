/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.util;

import static com.app.excelcsvconverter.consts.summary.SummaryConfig.CONFIG;

/**
 * 获取中英文信息
 *
 * @since 2019-08-17
 */
public class MessageUtil {

    public static String getMessage(String enKey, String cnKey, String language) {
        String message;
        if ("EN".equalsIgnoreCase(language)) {
            message = CONFIG.getConfig(enKey);
        } else {
            message = CONFIG.getConfig(cnKey);
        }
        return message;
    }
}
