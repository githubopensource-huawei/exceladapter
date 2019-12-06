/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.read;

import com.app.excelcsvconverter.JaxbContexts;
import com.app.excelcsvconverter.SaxSources;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.function.Function;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.Unmarshaller;
import javax.xml.parsers.ParserConfigurationException;

/**
 * 解析xml文件返回对象
 *
 * @since 2019-08-01
 */
public class JaxbReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(JaxbReader.class);

    public static <T> T readConfig(Class<T> clazz, File file) {
        return readConfig(clazz, Function.identity(), file);
    }

    private static <T> T readConfig(Class<T> clazz, Function<Unmarshaller, Unmarshaller> unmarshallerTweaker,
        File file) {
        if (file == null || !file.exists()) {
            return null;
        }
        try (InputStream inputStream = new FileInputStream(file)) {
            JAXBContext context = JaxbContexts.of(clazz);
            Unmarshaller unmarshaller = context.createUnmarshaller();
            unmarshaller = unmarshallerTweaker.apply(unmarshaller);
            T result = (T) unmarshaller.unmarshal(SaxSources.newSecurityUnmarshalSource(inputStream));
            return result;
        } catch (FileNotFoundException e) {
            LOGGER.error("File not found, file:{}", file.getName());
        } catch (IOException e) {
            LOGGER.error("Failed to readConfig file");
        } catch (JAXBException | ParserConfigurationException | SAXException ex) {
            LOGGER.error("Failed to unmarshall file:{}, class:{}", file.getName(), clazz);
        }
        return null;
    }
}
