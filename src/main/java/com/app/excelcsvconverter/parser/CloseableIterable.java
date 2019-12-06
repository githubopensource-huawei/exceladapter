/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.parser;

import java.io.Closeable;
import java.util.Iterator;
import java.util.List;

/**
 * 可关闭Iterable
 *
 * @since 2019-08-21
 */
public interface CloseableIterable extends Iterator<List<String>>, Closeable {

    @Override
    default void close() {
    }
}
