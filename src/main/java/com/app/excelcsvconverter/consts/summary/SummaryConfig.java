/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.consts.summary;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.Properties;

public enum SummaryConfig {
    CONFIG;

    private static final Logger LOGGER = LoggerFactory.getLogger(SummaryConfig.class);

    private static final String SUMMARY_CONFIG = "SummaryConfig.properties";

    private Properties configs;

    SummaryConfig() {
        init();
    }

    public synchronized void init() {
        InputStream resourceAsStream = SummaryConfig.class.getResourceAsStream(SUMMARY_CONFIG);
        configs = new Properties();
        try (InputStreamReader reader = new InputStreamReader(resourceAsStream, StandardCharsets.UTF_8)) {
            configs.load(reader);
        } catch (FileNotFoundException e) {
            LOGGER.error("File {} not found!", SUMMARY_CONFIG);
        } catch (IOException e) {
            LOGGER.error("Init SummaryMappings fail:{}", e);
        }
    }

    public String getConfig(String key) {
        return configs.getProperty(key, key);
    }

}