/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.writer;

import com.app.excelcsvconverter.JaxbContexts;
import com.app.excelcsvconverter.util.SecureFileUtil;
import com.google.common.base.Strings;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.OutputStream;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.Marshaller;

/**
 * 生成xml文件
 *
 * @since 2019-08-01
 */
public class JaxbWriter {

    private static final Logger LOGGER = LoggerFactory.getLogger(JaxbWriter.class);

    public static void writeConfig(Object jaxbElement, File config) {
        writeConfig(jaxbElement, config, null);
    }

    public static void writeConfig(Object jaxbElement, File config, String encode) {
        try (OutputStream fileOut = SecureFileUtil.getFileSafeOutputStream(config)) {
            LOGGER.info("write configFileName = {}", config);
            JAXBContext jc = JaxbContexts.of(jaxbElement.getClass());
            Marshaller m = jc.createMarshaller();
            if (!Strings.isNullOrEmpty(encode)) {
                LOGGER.info("write xml encode = {}", encode);
                m.setProperty(Marshaller.JAXB_ENCODING, encode);
            }
            m.setProperty(Marshaller.JAXB_FORMATTED_OUTPUT, true);
            m.marshal(jaxbElement, fileOut);
            fileOut.flush();
        } catch (Exception e) {
            LOGGER.info(config.getAbsolutePath(), e);
        }
    }
}
