/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXNotRecognizedException;
import org.xml.sax.SAXNotSupportedException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;

public final class SaxParserFactories {
    private static final Logger LOGGER = LoggerFactory.getLogger(SaxParserFactories.class);

    private SaxParserFactories() {
    }

    public static SAXParserFactory newSecurityInstance() {
        SAXParserFactory factory = SAXParserFactory.newInstance();

        try {
            factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        } catch (SAXNotRecognizedException | SAXNotSupportedException | ParserConfigurationException var3) {
            LOGGER.error("FAILED to set feature http://apache.org/xml/features/disallow-doctype-decl", var3);
        }

        try {
            factory.setFeature("http://javax.xml.XMLConstants/feature/secure-processing", true);
        } catch (Exception var2) {
            LOGGER.error("FAILED to set feature XMLConstants.FEATURE_SECURE_PROCESSING to true", var2);
        }

        factory.setValidating(true);
        return factory;
    }
}