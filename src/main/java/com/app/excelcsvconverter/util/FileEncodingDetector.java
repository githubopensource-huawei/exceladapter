/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * 功能描述：
 *
 * @since 2019-09-05
 */
public class FileEncodingDetector {
    protected static final String CHARSET_ASCII = "ASCII";

    protected static final String CHARSET_UNICODE_BIG = "UnicodeBig";

    protected static final String CHARSET_UNICODE_LITTLE = "UnicodeLittle";

    private static final Logger LOGGER = LoggerFactory.getLogger(FileEncodingDetector.class);

    private static final String CHARSET_NO_UTF_8 = "GBK";

    private static final String CHARSET_UTF_8 = "UTF-8";

    private static final char CHAR_EQUIVALENT_2_BITS = '\u0080';

    private static final char CHAR_EQUIVALENT_3_BITS = 'À';

    private static final char CHAR_EQUIVALENT_4_BITS = 'ÿ';

    protected static boolean havingBom = false;

    public FileEncodingDetector() {
    }

    public static String detectEncodingFromFile(File var0) throws IOException {
        String var3;
        try (FileInputStream var1 = new FileInputStream(var0)) {
            String var2 = detectEncodingFromStream(var1);
            LOGGER.debug("Encoding of file:{},is{}", var0.getName(), var2);
            var3 = var2;
        } catch (IOException e) {
            LOGGER.error("detect encoding fail.", e);
            throw e;
        }

        return var3;
    }

    public static String detect(File file) throws IOException {
        return detectEncodingFromFile(file);
    }

    private static String detectEncodingFromStream(InputStream var0) throws IOException {
        try (BufferedInputStream var1 = new BufferedInputStream(var0)) {
            var1.mark(0);
            havingBom = false;
            int var2 = var1.read();
            int var3 = var1.read();
            int var4 = var1.read();
            String var5 = Integer.toHexString(255 & var2);
            String var6 = Integer.toHexString(255 & var3);
            String var7 = Integer.toHexString(255 & var4);
            String var29;
            if ("FF".equalsIgnoreCase(var5) && "FE".equalsIgnoreCase(var6)) {
                var29 = "UnicodeLittle";
                return var29;
            } else if ("FE".equalsIgnoreCase(var5) && "FF".equalsIgnoreCase(var6)) {
                var29 = "UnicodeBig";
                return var29;
            } else if ("EF".equalsIgnoreCase(var5) && "BB".equalsIgnoreCase(var6) && "BF".equalsIgnoreCase(var7)) {
                havingBom = true;
                var29 = "UTF-8";
                return var29;
            } else {
                var1.reset();
                boolean var8 = false;
                boolean var9 = true;
                boolean var10 = true;
                int var27 = 0;

                String var11;
                int var28;
                while ((var28 = var1.read()) != -1) {
                    if ((var28 & 128) != 0) {
                        var10 = false;
                    }

                    if (var27 != 0) {
                        if ((var28 & 192) != 128) {
                            var11 = "GBK";
                            return var11;
                        }

                        --var27;
                    } else if (var28 >= 128) {
                        do {
                            var28 <<= 1;
                            ++var27;
                        } while ((var28 & 128) != 0);

                        --var27;
                        if (var27 == 0) {
                            var11 = "GBK";
                            return var11;
                        }
                    }
                }

                if (var27 > 0) {
                    var11 = "GBK";
                    return var11;
                } else if (var10) {
                    var11 = "ASCII";
                    return var11;
                } else {
                    var11 = "UTF-8";
                    return var11;
                }
            }
        } catch (IOException e) {
            LOGGER.error("detect encoding from stream fail.", e);
            throw e;
        }
    }
}
