/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.util;

import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.Locale;

public class EncodingUtil {
    public static Charset getDefaultEncoding() {
        String language = Locale.getDefault().getLanguage();
        if ("zh".equals(language)) {
            return Charset.forName("GBK");
        }

        if ("en".equals(language)) {
            return StandardCharsets.UTF_8;
        }
        return Charset.defaultCharset();
    }
}
