/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Closeable;
import java.io.IOException;
import java.util.zip.ZipFile;

public final class Closeables {
    private static final Logger LOGGER = LoggerFactory.getLogger(Closeables.class);

    private Closeables() {
    }

    public static void closeQuietly(ZipFile zipFile) {
        if (null != zipFile) {
            try {
                zipFile.close();
            } catch (IOException var2) {
                LOGGER.warn("Exception thrown while closing ZipFile.{}", zipFile.getName());
            }

        }
    }

    public static void closeQuietly(Closeable closeable) {
        if (null != closeable) {
            try {
                closeable.close();
            } catch (IOException var2) {
                LOGGER.warn("Exception thrown while closing Closeable.");
            }

        }
    }
}