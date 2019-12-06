/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter;

import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import java.io.InputStream;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import javax.xml.transform.sax.SAXSource;

public final class SaxSources {
    private SaxSources() {
    }

    public static SAXSource newSecurityUnmarshalSource(InputStream in)
        throws ParserConfigurationException, SAXException {
        SAXParserFactory factory = SaxParserFactories.newSecurityInstance();
        XMLReader xmlReader = factory.newSAXParser().getXMLReader();
        return new SAXSource(xmlReader, new InputSource(in));
    }
}