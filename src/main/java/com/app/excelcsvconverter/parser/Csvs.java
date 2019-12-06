/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.parser;

import java.io.IOException;
import java.io.StringWriter;
import java.io.Writer;

public final class Csvs {
    private static final char DEFAULT_DELIMITER = ',';

    private static final char DOUBLE_QUOTE = '"';

    private static final char CARRIAGE_RETURN = '\r';

    private static final char NEWLINE = '\n';

    public Csvs() {
    }

    public static String escape(String text) {
        return escape(text, ',');
    }

    public static String escape(String text, char delimiter) {
        try {
            StringWriter writer = new StringWriter();
            escape(writer, text, delimiter);
            return writer.toString();
        } catch (IOException var3) {
            throw new RuntimeException("escapeCSV error", var3);
        }
    }

    public static void escape(Writer writer, String text, char delimiter) throws IOException {
        StringBuilder sbuf = new StringBuilder();
        boolean buffering = true;

        for (int index = 0; index < text.length(); ++index) {
            char ch = text.charAt(index);
            switch (ch) {
                case '\n':
                case '\r':
                    if (buffering) {
                        buffering = false;
                        writer.write(34);
                        if (sbuf.length() != 0) {
                            writer.write(sbuf.toString());
                        }
                    }

                    writer.write(ch);
                    break;
                case '"':
                    if (buffering) {
                        buffering = false;
                        writer.write(34);
                        if (sbuf.length() != 0) {
                            writer.write(sbuf.toString());
                        }
                    }

                    writer.write(34);
                    writer.write(34);
                    break;
                default:
                    if (ch == delimiter) {
                        if (buffering) {
                            buffering = false;
                            writer.write(34);
                            if (0 != sbuf.length()) {
                                writer.write(sbuf.toString());
                            }
                        }

                        writer.write(ch);
                    } else if (buffering) {
                        sbuf.append(ch);
                    } else {
                        writer.write(ch);
                    }
            }
        }

        if (buffering) {
            writer.write(text);
        } else {
            writer.write(34);
        }
    }
}
