/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.parser;

import com.app.excelcsvconverter.Closeables;
import com.app.excelcsvconverter.util.SecureFileUtil;
import com.google.common.io.Files;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BlankRecord;
import org.apache.poi.hssf.record.BoolErrRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.FormulaRecord;
import org.apache.poi.hssf.record.LabelRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.hssf.record.StringRecord;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Xls2Csv implements HSSFListener {

    private static final Logger LOGGER = LoggerFactory.getLogger(Xls2Csv.class);

    private POIFSFileSystem fs;

    private BufferedWriter output;

    private int lastRowNumber;

    private int lastColumnNumber;

    private boolean outputFormulaValues;

    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;

    private HSSFWorkbook stubWorkbook;

    private SSTRecord sstRecord;

    private FormatTrackingHSSFListener formatListener;

    private int sheetIndex;

    private BoundSheetRecord[] orderedBSRs;

    private List<BoundSheetRecord> boundSheetRecords;

    private int nextRow;

    private int nextColumn;

    private boolean outputNextStringRecord;

    private boolean writeVeryHiddenSheet;

    private boolean writeSheetState;

    private File outDir;

    private Charset outEncoding;

    private List<String> sheetNameList;

    private boolean trim;

    private String sheetName;

    private StringBuilder lastRowBuilder;

    private boolean firstWrite;

    private Map<String, Integer> sheetStatusMap = new HashMap<>();

    public Xls2Csv(File file, File outDir, Charset outEncoding, List<String> sheetNameList, boolean writeVeryHiddenSheet, boolean writeSheetState) throws IOException {
        this(file, outDir, outEncoding, sheetNameList, writeVeryHiddenSheet, writeSheetState, true);
    }

    public Xls2Csv(File file, File outDir, Charset outEncoding, List<String> sheetNameList, boolean writeVeryHiddenSheet, boolean writeSheetState, boolean trim) throws IOException {
        this.lastRowNumber = -1;
        this.lastColumnNumber = -1;
        this.outputFormulaValues = true;
        this.sheetIndex = -1;
        this.boundSheetRecords = new ArrayList();
        this.lastRowBuilder = new StringBuilder();
        this.firstWrite = false;
        InputStream inputStream = null;

        try {
            inputStream = SecureFileUtil.getFileSafeInputStream(file, false);
            this.fs = new POIFSFileSystem(inputStream);
            this.writeVeryHiddenSheet = writeVeryHiddenSheet;
            this.writeSheetState = writeSheetState;
            this.sheetNameList = sheetNameList;
        } finally {
            Closeables.closeQuietly(inputStream);
        }

        this.outDir = outDir;
        this.outEncoding = outEncoding;
        this.trim = trim;
    }

    public void process() throws IOException {
        HSSFRequest request = this.getHssfRequest();
        HSSFEventFactory factory = new HSSFEventFactory();
        factory.processWorkbookEvents(request, this.fs);
        if (this.writeSheetState) {
            this.writeSheetState();
        }
        this.flushAndCloseOutput();
    }

    private HSSFRequest getHssfRequest() {
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        this.formatListener = new FormatTrackingHSSFListener(listener);
        HSSFRequest request = new HSSFRequest();
        if (this.outputFormulaValues) {
            request.addListenerForAllRecords(this.formatListener);
        } else {
            this.workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(this.formatListener);
            request.addListenerForAllRecords(this.workbookBuildingListener);
        }

        return request;
    }

    private void flushAndCloseOutput() throws IOException {
        if (null != this.output) {
            this.writeRowToFile(this.output, this.lastRowBuilder);
            this.output.flush();
            this.output.close();
            this.output = null;
        }

    }

    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;

        try {
            switch (record.getSid()) {
                case 6:
                    if (null != this.output) {
                        FormulaRecord frec = (FormulaRecord) record;
                        thisRow = frec.getRow();
                        thisColumn = frec.getColumn();
                        if (this.outputFormulaValues) {
                            this.outputNextStringRecord = true;
                            this.nextRow = frec.getRow();
                            this.nextColumn = frec.getColumn();
                        } else {
                            thisStr = '"' + HSSFFormulaParser.toFormulaString(this.stubWorkbook, frec.getParsedExpression()) + '"';
                        }
                    }
                    break;
                case 133:
                    this.boundSheetRecords.add((BoundSheetRecord) record);
                    break;
                case 252:
                    this.sstRecord = (SSTRecord) record;
                    break;
                case 253:
                    if (null != this.output) {
                        LabelSSTRecord lsrec = (LabelSSTRecord) record;
                        thisRow = lsrec.getRow();
                        thisColumn = lsrec.getColumn();
                        thisStr = this.sstRecord == null
                            ? "\"(No SST Record, can't identify string)\""
                            : this.sstRecord.getString(lsrec.getSSTIndex()).toString();
                    }
                    break;
                case 513:
                    if (null != this.output) {
                        BlankRecord brec = (BlankRecord) record;
                        thisRow = brec.getRow();
                        thisColumn = brec.getColumn();
                        thisStr = "";
                    }
                    break;
                case 515:
                    if (null != this.output) {
                        NumberRecord numrec = (NumberRecord) record;
                        thisRow = numrec.getRow();
                        thisColumn = numrec.getColumn();
                        thisStr = this.formatListener.formatNumberDateCell(numrec);
                    }
                    break;
                case 516:
                    if (null != this.output) {
                        LabelRecord lrec = (LabelRecord) record;
                        thisRow = lrec.getRow();
                        thisColumn = lrec.getColumn();
                        thisStr = lrec.getValue();
                    }
                    break;
                case 517:
                    if (null != this.output) {
                        BoolErrRecord berec = (BoolErrRecord) record;
                        thisRow = berec.getRow();
                        thisColumn = berec.getColumn();
                        thisStr = "";
                    }
                    break;
                case 519:
                    if (this.outputNextStringRecord && null != this.output) {
                        StringRecord srec = (StringRecord) record;
                        thisStr = srec.getString();
                        thisRow = this.nextRow;
                        thisColumn = this.nextColumn;
                        this.outputNextStringRecord = false;
                    }
                    break;
                case 2057:
                    this.processBofRecord((BOFRecord) record);
            }

            if (null != this.output) {
                boolean needNewRow = false;
                int i;
                if (thisRow > this.lastRowNumber && thisRow != this.lastRowNumber + 1) {
                    if (this.lastRowNumber != -1) {
                        this.writeRowToFile(this.output, this.lastRowBuilder);
                    }

                    for (i = this.lastRowNumber + 1; i < thisRow; ++i) {
                        for (i = 0; i <= this.lastColumnNumber; ++i) {
                            this.lastRowBuilder.append(",");
                        }

                        this.writeRowToFile(this.output, this.lastRowBuilder);
                        ++this.lastRowNumber;
                    }

                    this.lastColumnNumber = 0;
                    this.firstWrite = false;
                    needNewRow = true;
                } else if (thisRow != -1 && thisRow != this.lastRowNumber) {
                    if (this.lastRowNumber != -1) {
                        this.writeRowToFile(this.output, this.lastRowBuilder);
                    }

                    this.lastColumnNumber = 0;
                    needNewRow = true;
                }

                if (thisStr != null) {
                    boolean needTrim = ("Cover".equals(this.sheetName) || "封面".equals(this.sheetName)) && this.trim;
                    if ((!needNewRow || !needTrim) && thisColumn > 0) {
                        for (i = this.lastColumnNumber; i < thisColumn; ++i) {
                            this.lastRowBuilder.append(',');
                        }
                    }

                    String str = Csvs.escape(thisStr);
                    this.lastRowBuilder.append(str);
                }
            }

            if (thisRow > -1) {
                this.lastRowNumber = thisRow;
            }

            if (thisColumn > -1) {
                this.lastColumnNumber = thisColumn;
            }
        } catch (FileNotFoundException var8) {
            LOGGER.error("The output file not found");
        } catch (IOException var9) {
            LOGGER.error(var9.getMessage(), var9);
        }

    }

    private void processBofRecord(BOFRecord record) throws IOException {
        if (record.getType() == 16) {
            if (this.workbookBuildingListener != null && this.stubWorkbook == null) {
                this.stubWorkbook = this.workbookBuildingListener.getStubHSSFWorkbook();
            }

            ++this.sheetIndex;
            if (this.orderedBSRs == null) {
                this.orderedBSRs = BoundSheetRecord.orderByBofPosition(this.boundSheetRecords);
                for (BoundSheetRecord orderedBSR : orderedBSRs) {
                    sheetStatusMap.put(orderedBSR.getSheetname(), getVisible(orderedBSR));
                }
            }

            this.sheetName = this.orderedBSRs[this.sheetIndex].getSheetname();
            this.flushAndCloseOutput();
            BoundSheetRecord sheetRec = this.orderedBSRs[this.sheetIndex];
            this.initCsvFile(sheetRec);
        }

    }

    private void initCsvFile(BoundSheetRecord sheetRec) throws FileNotFoundException {
        boolean valid = false;
        if (!this.sheetNameList.isEmpty()) {
            if (this.sheetNameList.contains(this.sheetName)) {
                valid = true;
            }
        } else if (this.writeVeryHiddenSheet || !sheetRec.isVeryHidden()) {
            valid = true;
        }

        if (valid) {
            File file;
            if (sheetStatusMap.get(this.sheetName) != -1) {
                file = new File(this.outDir, "\\Internal");
                if (!file.isDirectory()) {
                    file.mkdir();
                }
                file = new File(file, this.sheetName + ".csv");
            } else {
                file = new File(this.outDir, this.sheetName + ".csv");
            }
            this.output = new BufferedWriter(Files.newWriter(file, this.outEncoding));
            this.firstWrite = true;
            this.lastRowNumber = -1;
            this.lastColumnNumber = -1;
        }

    }

    private void writeSheetState() {
        File file = new File(this.outDir, "SHEET STYLE.csv");

        try {
            BufferedWriter output = new BufferedWriter(Files.newWriter(file, this.outEncoding));
            Throwable var3 = null;

            try {
                StringBuilder rowBuilder = new StringBuilder();
                boolean isFirst = true;
                BoundSheetRecord[] var6 = this.orderedBSRs;
                int var7 = var6.length;

                for (BoundSheetRecord ordered : var6) {
                    if (isFirst) {
                        isFirst = false;
                    } else {
                        output.newLine();
                    }

                    int visible = -1;
                    if (ordered.isHidden()) {
                        visible = 0;
                    } else if (ordered.isVeryHidden()) {
                        visible = 2;
                    }

                    rowBuilder.append(ordered.getSheetname()).append(",").append(visible);
                    output.write(rowBuilder.toString());
                    rowBuilder.setLength(0);
                }
                output.flush();
            } catch (Throwable var20) {
                var3 = var20;
                throw var20;
            } finally {
                if (output != null) {
                    if (var3 != null) {
                        try {
                            output.close();
                        } catch (Throwable var19) {
                            var3.addSuppressed(var19);
                        }
                    } else {
                        output.close();
                    }
                }

            }
        } catch (FileNotFoundException var22) {
            LOGGER.error("{} not found", file.getName());
        } catch (IOException var23) {
            LOGGER.error(var23.getMessage(), var23);
        }

    }

    private void writeRowToFile(BufferedWriter output, StringBuilder rowBuilder) throws IOException {
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

    private int getVisible(BoundSheetRecord ordered) {
        int visible;
        if (ordered.isHidden()) {
            visible = 0;
        } else if (ordered.isVeryHidden()) {
            visible = 2;
        } else {
            visible = -1;
        }
        return visible;
    }
}
