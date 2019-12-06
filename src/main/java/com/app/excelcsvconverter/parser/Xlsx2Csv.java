/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.parser;

import com.app.excelcsvconverter.SaxParserFactories;
import com.google.common.base.Strings;
import com.google.common.collect.Maps;
import com.google.common.io.Files;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STSheetState;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.WorkbookDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;

public class Xlsx2Csv {
    private static final Logger LOGGER = LoggerFactory.getLogger(Xlsx2Csv.class);

    private static final String INLINE_STR = "inlineStr";

    private static final String TAG_V = "v";

    private static final String TAG_T = "t";

    private static final String TAG_R = "r";

    private static final String TAG_S = "s";

    private static final String TAG_B = "b";

    private static final String TAG_E = "e";

    private static final String TAG_STR = "str";

    private static final String TAG_N = "n";

    private static final Pattern START_PATTERN = Pattern.compile("^\\s+");

    private static final Pattern END_PATTERN = Pattern.compile("\\s+$");

    private boolean trim;

    private boolean writeVeryHiddenSheet;

    private boolean writeSheetState;

    private OPCPackage xlsxPackage;

    private File outDir;

    private Charset outEncoding;

    private BufferedWriter output;

    private List<String> sheetNameList;

    private boolean needTrim;

    private StringBuilder lastRowBuilder;

    private Map<String, Integer> statusMap = new HashMap<>();

    public Xlsx2Csv(OPCPackage pkg, File outDir, Charset outEncoding, List<String> sheetNameList, boolean writeVeryHiddenSheet, boolean writeSheetState) {
        this(pkg, outDir, outEncoding, sheetNameList, writeVeryHiddenSheet, writeSheetState, true);
    }

    public Xlsx2Csv(OPCPackage pkg, File outDir, Charset outEncoding, List<String> sheetNameList, boolean writeVeryHiddenSheet, boolean writeSheetState, boolean trim) {
        this.lastRowBuilder = new StringBuilder();
        this.xlsxPackage = pkg;
        this.outDir = outDir;
        this.outEncoding = outEncoding;
        this.sheetNameList = sheetNameList;
        this.writeVeryHiddenSheet = writeVeryHiddenSheet;
        this.writeSheetState = writeSheetState;
        this.trim = trim;
    }

    public void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings, InputStream sheetInputStream)
        throws IOException, ParserConfigurationException, SAXException {
        SAXParserFactory saxParserFactory = SaxParserFactories.newSecurityInstance();
        saxParserFactory.setValidating(false);
        XMLReader sheetParser = saxParserFactory.newSAXParser().getXMLReader();
        Xlsx2Csv.MyXSSFSheetHandler handler = new Xlsx2Csv.MyXSSFSheetHandler(styles, strings, this.output);
        sheetParser.setContentHandler(handler);
        InputSource sheetSource = new InputSource(sheetInputStream);
        sheetParser.parse(sheetSource);
        handler.writeRowToFile(this.output, this.lastRowBuilder);
    }

    public void process() throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator sheetIter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        InputStream workbookData = xssfReader.getWorkbookData();
        Map needWriteCsvSheetMap = this.getValidSheetMapAndWriteSheetState(workbookData);

        while (sheetIter.hasNext()) {
            InputStream stream = sheetIter.next();
            String sheetName = sheetIter.getSheetName();
            this.needTrim = ("Cover".equals(sheetName) || "封面".equals(sheetName)) && this.trim;
            if ((Boolean) needWriteCsvSheetMap.get(sheetName)) {
                this.flushAndCloseOutput();
                File file;
                if (statusMap.get(sheetName) != -1) {
                    file = new File(this.outDir, "\\Internal");
                    if (!file.isDirectory()) {
                        file.mkdir();
                    }
                    file = new File(file, sheetName + ".csv");
                } else {
                    file = new File(this.outDir, sheetName + ".csv");
                }
                this.output = new BufferedWriter(Files.newWriter(file, this.outEncoding));
                this.processSheet(styles, strings, stream);
                stream.close();
            }
        }

        this.flushAndCloseOutput();
    }

    private Map<String, Boolean> getValidSheetMapAndWriteSheetState(InputStream workbookData) throws IOException {
        HashMap needWriteCsvSheetMap = Maps.newHashMap();

        try {
            WorkbookDocument document = WorkbookDocument.Factory.parse(workbookData);
            CTSheet[] var4 = document.getWorkbook().getSheets().getSheetArray();
            int var5 = var4.length;

            for (CTSheet ctSheet : var4) {
                needWriteCsvSheetMap.put(ctSheet.getName(), this.isValidSheet(ctSheet));
            }

            if (this.writeSheetState) {
                this.writeSheetState(document);
            }
        } catch (XmlException var8) {
            LOGGER.error(var8.getMessage(), var8);
        }

        return needWriteCsvSheetMap;
    }

    private boolean isValidSheet(CTSheet ctSheet) {
        return !this.sheetNameList.isEmpty()
            ? this.sheetNameList.contains(ctSheet.getName())
            : this.writeVeryHiddenSheet || !STSheetState.VERY_HIDDEN.equals(ctSheet.getState());
    }

    private void writeSheetState(WorkbookDocument document) {
        File file = new File(this.outDir, "SHEET STYLE.csv");

        try {
            BufferedWriter output = new BufferedWriter(Files.newWriter(file, this.outEncoding));
            Throwable var4 = null;

            try {
                StringBuilder rowBuilder = new StringBuilder();
                boolean isFirst = true;
                CTSheet[] var7 = document.getWorkbook().getSheets().getSheetArray();
                int var8 = var7.length;

                for (CTSheet ctSheet : var7) {
                    if (isFirst) {
                        isFirst = false;
                    } else {
                        output.newLine();
                    }

                    int visible = -1;
                    if (ctSheet.getState() == STSheetState.HIDDEN) {
                        visible = 0;
                    } else if (ctSheet.getState() == STSheetState.VERY_HIDDEN) {
                        visible = 2;
                    }
                    statusMap.put(ctSheet.getName(), visible);
                    rowBuilder.append(ctSheet.getName()).append(",").append(visible);
                    output.write(rowBuilder.toString());
                    rowBuilder.setLength(0);
                }

                output.flush();
            } catch (Throwable var20) {
                var4 = var20;
                throw var20;
            } finally {
                if (output != null) {
                    if (var4 != null) {
                        try {
                            output.close();
                        } catch (Throwable var19) {
                            var4.addSuppressed(var19);
                        }
                    } else {
                        output.close();
                    }
                }

            }
        } catch (IOException var22) {
            LOGGER.debug(var22.getMessage(), var22);
        }

    }

    private void flushAndCloseOutput() throws IOException {
        if (null != this.output) {
            this.output.flush();
            this.output.close();
            this.output = null;
        }

    }

    enum XssfDataType {
        BOOL,
        ERROR,
        FORMULA,
        INLINESTR,
        SSTINDEX,
        NUMBER;

        XssfDataType() {
        }
    }

    class MyXSSFSheetHandler extends DefaultHandler {
        private final BufferedWriter output;

        private final DataFormatter formatter;

        private StylesTable stylesTable;

        private ReadOnlySharedStringsTable sharedStringsTable;

        private boolean vIsOpen;

        private Xlsx2Csv.XssfDataType nextDataType;

        private short formatIndex;

        private String formatString;

        private int thisRow = -1;

        private int lastRowNumber = -1;

        private int thisColumn = -1;

        private int lastColumnNumber = -1;

        private StringBuilder value;

        private boolean isCellValueEmpty = false;

        private XSSFRichTextString sharingRichTextString = new XSSFRichTextString();

        private boolean firstWrite = true;

        public MyXSSFSheetHandler(StylesTable styles, ReadOnlySharedStringsTable strings, BufferedWriter target) {
            this.stylesTable = styles;
            this.sharedStringsTable = strings;
            this.output = target;
            this.value = new StringBuilder();
            this.nextDataType = Xlsx2Csv.XssfDataType.NUMBER;
            this.formatter = new DataFormatter();
        }

        public void startElement(String uri, String localName, String name, Attributes attributes) {
            if (this.isOpenTag(name)) {
                this.vIsOpen = true;
                this.value.setLength(0);
            } else if ("c".equals(name)) {
                this.processCTag(attributes);
            }
        }

        private boolean isOpenTag(String name) {
            return "inlineStr".equals(name) || "v".equals(name) || "t".equals(name);
        }

        private void processCTag(Attributes attributes) {
            String r = attributes.getValue("r");
            int firstDigit = this.getFirstDigit(r);
            this.thisColumn = this.nameToColumn(r.substring(0, firstDigit));
            this.thisRow = Integer.parseInt(r.substring(firstDigit));
            this.nextDataType = Xlsx2Csv.XssfDataType.NUMBER;
            this.formatIndex = -1;
            this.formatString = null;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if (null == cellType || "n".equals(cellType) || INLINE_STR.equals(cellType)) {
                this.isCellValueEmpty = true;
            }

            if ("b".equals(cellType)) {
                this.nextDataType = Xlsx2Csv.XssfDataType.BOOL;
            } else if ("e".equals(cellType)) {
                this.nextDataType = Xlsx2Csv.XssfDataType.ERROR;
            } else if ("inlineStr".equals(cellType)) {
                this.nextDataType = Xlsx2Csv.XssfDataType.INLINESTR;
            } else if ("s".equals(cellType)) {
                this.nextDataType = Xlsx2Csv.XssfDataType.SSTINDEX;
            } else if ("str".equals(cellType)) {
                this.nextDataType = Xlsx2Csv.XssfDataType.FORMULA;
            } else if (cellStyleStr != null) {
                this.processFormatCell(cellStyleStr);
            }

        }

        private void processFormatCell(String cellStyleStr) {
            int styleIndex = Integer.parseInt(cellStyleStr);
            XSSFCellStyle style = this.stylesTable.getStyleAt(styleIndex);
            if (style != null) {
                this.formatIndex = style.getDataFormat();
                this.formatString = style.getDataFormatString();
                if (this.formatString == null) {
                    this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                }
            }

        }

        private int getFirstDigit(String r) {
            int firstDigit = -1;

            for (int c = 0; c < r.length(); ++c) {
                if (Character.isDigit(r.charAt(c))) {
                    firstDigit = c;
                    break;
                }
            }

            return firstDigit;
        }

        public void endElement(String uri, String localName, String name) {
            String thisStr = null;

            try {
                String str;
                int col;
                if (!"v".equals(name) && !"t".equals(name)) {
                    if (this.isCellValueEmpty) {
                        thisStr = "";
                        this.isCellValueEmpty = false;
                    } else if (Objects.equals(this.nextDataType, Xlsx2Csv.XssfDataType.SSTINDEX)) {
                        thisStr = "";
                    }
                } else {
                    switch (nextDataType) {
                        case BOOL:
                            char first = this.value.charAt(0);
                            thisStr = first == 48 ? "FALSE" : "TRUE";
                            break;
                        case ERROR:
                            thisStr = "\"ERROR:" + this.value.toString() + '"';
                            break;
                        case FORMULA:
                            thisStr = this.value.toString();
                            break;
                        case INLINESTR:
                            thisStr = this.getDisplayValue(this.value.toString());
                            break;
                        case SSTINDEX:
                            str = this.value.toString();

                            try {
                                col = Integer.parseInt(str);
                                thisStr = this.getDisplayValue(this.sharedStringsTable.getEntryAt(col));
                            } catch (NumberFormatException var8) {
                                this.output.newLine();
                                this.output.write("Failed to parse SST index '" + str + "': " + var8.toString());
                            }
                            break;
                        case NUMBER:
                            String n = this.value.toString();
                            thisStr = this.formatString != null ? this.formatter.formatRawCellContents(
                                Double.parseDouble(n), this.formatIndex, this.formatString) : n;
                            break;
                        default:
                            thisStr = "(TODO: Unexpected type: " + this.nextDataType + ")";
                    }
                }

                if (null != thisStr) {
                    if (this.lastColumnNumber == -1) {
                        this.lastColumnNumber = 0;
                    }

                    boolean needNewRow = false;
                    if (this.lastRowNumber == -1) {
                        this.lastRowNumber = 0;
                    }

                    int i;
                    if (this.thisRow > this.lastRowNumber && this.thisRow != this.lastRowNumber + 1) {
                        if (this.lastRowNumber != 0) {
                            this.writeRowToFile(this.output, Xlsx2Csv.this.lastRowBuilder);
                        }

                        for (i = this.lastRowNumber + 1; i < this.thisRow; ++i) {
                            for (col = 0; col <= this.lastColumnNumber; ++col) {
                                Xlsx2Csv.this.lastRowBuilder.append(",");
                            }

                            this.writeRowToFile(this.output, Xlsx2Csv.this.lastRowBuilder);
                            ++this.lastRowNumber;
                        }

                        this.lastColumnNumber = 0;
                        this.firstWrite = false;
                        needNewRow = true;
                    } else if (this.thisRow != -1 && this.thisRow != this.lastRowNumber) {
                        if (this.lastRowNumber != 0) {
                            this.writeRowToFile(this.output, Xlsx2Csv.this.lastRowBuilder);
                        }

                        this.lastColumnNumber = 0;
                        needNewRow = true;
                    }

                    if ((!needNewRow || !Xlsx2Csv.this.needTrim) && this.thisColumn > 0) {
                        for (i = this.lastColumnNumber; i < this.thisColumn; ++i) {
                            Xlsx2Csv.this.lastRowBuilder.append(',');
                        }
                    }

                    str = Csvs.escape(thisStr);
                    Xlsx2Csv.this.lastRowBuilder.append(str);
                }

                if (this.thisColumn > -1) {
                    this.lastColumnNumber = this.thisColumn;
                }

                if (this.thisRow > -1) {
                    this.lastRowNumber = this.thisRow;
                }
            } catch (FileNotFoundException var9) {
                Xlsx2Csv.LOGGER.error("The output file not found");
            } catch (IOException var10) {
                Xlsx2Csv.LOGGER.error(var10.getMessage(), var10);
            } catch (Exception e) {
                Xlsx2Csv.LOGGER.error(e.getMessage());
            }

        }

        private String getDisplayValue(String str) {
            if (Strings.isNullOrEmpty(str)) {
                return str;
            } else {
                boolean supportRichText = true;
                if (supportRichText) {
                    String whitespace = " ";
                    if (!str.endsWith(whitespace) && !str.startsWith(whitespace)) {
                        this.sharingRichTextString.setString(str);
                        return this.sharingRichTextString.toString();
                    } else {
                        int startWhitespaceCount = 0;
                        Matcher startMatcher = Xlsx2Csv.START_PATTERN.matcher(str);
                        if (startMatcher.find()) {
                            startWhitespaceCount = startMatcher.group().length();
                        }

                        Matcher endMatcher = Xlsx2Csv.END_PATTERN.matcher(str);
                        int endWhitespaceCount = 0;
                        if (endMatcher.find()) {
                            endWhitespaceCount = endMatcher.group().length();
                        }

                        this.sharingRichTextString.setString(str.trim());
                        String result = this.sharingRichTextString.toString();
                        return this.fillWhitespace(result, startWhitespaceCount, endWhitespaceCount);
                    }
                } else {
                    return str;
                }
            }
        }

        private String fillWhitespace(String str, int startWhitespaceCount, int endWhitespaceCount) {
            String result = str;
            if (startWhitespaceCount > 0) {
                result = this.padLeft(str, str.length() + startWhitespaceCount);
            }

            if (endWhitespaceCount > 0) {
                result = this.padRight(result, result.length() + endWhitespaceCount);
            }

            return result;
        }

        private String padRight(String str, int length) {
            return String.format("%1$-" + length + "s", str);
        }

        private String padLeft(String str, int length) {
            return String.format("%1$" + length + "s", str);
        }

        void writeRowToFile(BufferedWriter output, StringBuilder rowBuilder) throws IOException {
            try {
                if (!this.firstWrite) {
                    output.newLine();
                }

                if (rowBuilder.length() == 0) {
                    rowBuilder.append(",");
                }

                output.write(rowBuilder.toString());
            } finally {
                this.firstWrite = false;
                rowBuilder.setLength(0);
            }

        }

        public void characters(char[] ch, int start, int length) {
            if (this.vIsOpen) {
                this.value.append(ch, start, length);
            }

        }

        private int nameToColumn(String name) {
            int column = -1;

            for (int i = 0; i < name.length(); ++i) {
                int c = name.charAt(i);
                column = (column + 1) * 26 + c - 65;
            }

            return column;
        }
    }
}

