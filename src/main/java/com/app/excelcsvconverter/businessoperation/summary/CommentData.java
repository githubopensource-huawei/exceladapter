/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.businessoperation.summary;

/**
 * CommentData实体
 *
 * @since 2019-08-08
 */
public class CommentData {
    private int row;

    private short column;

    private String message;

    public CommentData(int row, short column, String message) {
        this.row = row;
        this.column = column;
        this.message = message;
    }

    public int getRow() {
        return this.row;
    }

    public short getColumn() {
        return this.column;
    }

    public String getMessage() {
        return this.message;
    }
}