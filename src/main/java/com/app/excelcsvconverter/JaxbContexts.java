/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter;

import com.google.common.base.Stopwatch;
import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;
import com.google.common.io.ByteSource;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.concurrent.TimeUnit;
import java.util.function.Function;
import java.util.function.Supplier;

import javax.xml.bind.JAXBContext;
import javax.xml.bind.JAXBException;
import javax.xml.bind.PropertyException;
import javax.xml.bind.Unmarshaller;
import javax.xml.parsers.ParserConfigurationException;

public final class JaxbContexts {
    private static final Logger LOGGER = LoggerFactory.getLogger(JaxbContexts.class);

    private static final LoadingCache<Class<?>, JAXBContext> CLASS_CONTEXT_CACHE;

    static {
        CacheBuilder var10000 = CacheBuilder.newBuilder().softValues();
        JaxbContexts.ClassContextCreator var10001 = JaxbContexts.ClassContextCreator.SINGLETON;
        JaxbContexts.ClassContextCreator.SINGLETON.getClass();
        CLASS_CONTEXT_CACHE = var10000.build(CacheLoader.from(var10001::apply));
    }

    private JaxbContexts() {
    }

    public static <T> Supplier<T> instanceSupplier(Class<T> clazz, File file) {
        return instanceSupplier(clazz, Function.identity(), file);
    }

    public static <T, F> Supplier<T> instanceSupplier(Class<T> clazz, F objectFactory, File file) {
        return new JaxbContexts.FileInstanceSupplier(clazz, newObjectFactoryUnmarshallerTweaker(objectFactory), file);
    }

    public static <T> Supplier<T> instanceSupplier(Class<T> clazz,
        Function<Unmarshaller, Unmarshaller> unmarshallerTweaker, File file) {
        return new JaxbContexts.FileInstanceSupplier(clazz, unmarshallerTweaker, file);
    }

    public static <T> Supplier<T> instanceSupplier(Class<T> clazz, InputStream inputStream) {
        return instanceSupplier(clazz, Function.identity(), inputStream);
    }

    public static <T> Supplier<T> instanceSupplier(Class<T> clazz,
        Function<Unmarshaller, Unmarshaller> unmarshallerTweaker, InputStream inputStream) {
        return new JaxbContexts.FileInstanceSupplier(clazz, unmarshallerTweaker, inputStream);
    }

    public static <T> Supplier<T> instanceSupplier(Class<T> clazz, ByteSource byteSource) {
        return instanceSupplier(clazz, Function.identity(), byteSource);
    }

    public static <T, F> Supplier<T> instanceSupplier(Class<T> clazz, F objectFactory, ByteSource byteSource) {
        return new JaxbContexts.InputStreamInstanceSupplier(clazz, newObjectFactoryUnmarshallerTweaker(objectFactory),
            byteSource);
    }

    public static <T> Supplier<T> instanceSupplier(Class<T> clazz,
        Function<Unmarshaller, Unmarshaller> unmarshallerTweaker, ByteSource byteSource) {
        return new JaxbContexts.InputStreamInstanceSupplier(clazz, unmarshallerTweaker, byteSource);
    }

    public static ThreadLocal<Unmarshaller> newThreadLocalUnmarshaller(Class<?> clazz) {
        return new JaxbContexts.ThreadLocalUnmarshaller(clazz);
    }

    public static <T> Function<Unmarshaller, Unmarshaller> newObjectFactoryUnmarshallerTweaker(T objectFactory) {
        return new JaxbContexts.ObjectFactoryUnmarshallerTweaker(objectFactory);
    }

    public static void reset() {
        CLASS_CONTEXT_CACHE.cleanUp();
    }

    public static JAXBContext of(Class<?> clazz) {
        return CLASS_CONTEXT_CACHE.getUnchecked(clazz);
    }

    private enum ClassContextCreator implements Function<Class<?>, JAXBContext> {
        SINGLETON;

        ClassContextCreator() {
        }

        public JAXBContext apply(Class<?> clazz) {
            Stopwatch stopwatch = Stopwatch.createStarted();

            JAXBContext var3;
            try {
                var3 = JAXBContext.newInstance(clazz);
            } catch (JAXBException var7) {
                throw new IllegalStateException("FAILED to resolve jaxb context for class " + clazz, var7);
            } finally {
                JaxbContexts.LOGGER.debug("obtain jaxb context cost: {}, {}", stopwatch.elapsed(TimeUnit.MILLISECONDS),
                    clazz);
            }

            return var3;
        }
    }

    private static class ObjectFactoryUnmarshallerTweaker<T> implements Function<Unmarshaller, Unmarshaller> {
        private T objectFactory;

        public ObjectFactoryUnmarshallerTweaker(T objectFactory) {
            this.objectFactory = objectFactory;
        }

        public Unmarshaller apply(Unmarshaller unmarshaller) {
            try {
                unmarshaller.setProperty("com.sun.xml.internal.bind.ObjectFactory", this.objectFactory);
            } catch (PropertyException var3) {
                JaxbContexts.LOGGER.error("FAILED to set property: com.sun.xml.internal.bind.ObjectFactory", var3);
            }

            return unmarshaller;
        }
    }

    private static class ThreadLocalUnmarshaller extends ThreadLocal<Unmarshaller> {
        private Class<?> clazz;

        public ThreadLocalUnmarshaller(Class<?> clazz) {
            this.clazz = clazz;
        }

        protected Unmarshaller initialValue() {
            try {
                return JaxbContexts.of(this.clazz).createUnmarshaller();
            } catch (JAXBException var2) {
                throw new IllegalStateException("FAILED to create unmarshaller for class " + this.clazz, var2);
            }
        }
    }

    private static class InputStreamInstanceSupplier<T> implements Supplier<T> {
        private Class<T> clazz;

        private Function<Unmarshaller, Unmarshaller> unmarshallerTweaker;

        private ByteSource byteSource;

        public InputStreamInstanceSupplier(Class<T> clazz, Function<Unmarshaller, Unmarshaller> unmarshallerTweaker,
            ByteSource byteSource) {
            this.clazz = clazz;
            this.unmarshallerTweaker = unmarshallerTweaker;
            this.byteSource = byteSource;
        }

        public T get() {
            Stopwatch stopwatch = Stopwatch.createStarted();
            InputStream in = null;

            try {
                in = this.byteSource.openBufferedStream();
                JAXBContext context = JaxbContexts.of(this.clazz);
                Unmarshaller unmarshaller = context.createUnmarshaller();
                unmarshaller = this.unmarshallerTweaker.apply(unmarshaller);
                T result = (T) unmarshaller.unmarshal(SaxSources.newSecurityUnmarshalSource(in));
                T var6 = result;
                return var6;
            } catch (Exception var10) {
                JaxbContexts.LOGGER.error("FAILED to unmarshall class:" + this.clazz, var10);
            } finally {
                Closeables.closeQuietly(in);
                JaxbContexts.LOGGER.debug("unmarshall cost: {}, clazz: {}", stopwatch.elapsed(TimeUnit.MILLISECONDS),
                    this.clazz);
            }

            return null;
        }
    }

    private static class FileInstanceSupplier<T> implements Supplier<T> {
        private Class<T> clazz;

        private Function<Unmarshaller, Unmarshaller> unmarshallerTweaker;

        private File file;

        private InputStream inputStream;

        public FileInstanceSupplier(Class<T> clazz, Function<Unmarshaller, Unmarshaller> unmarshallerTweaker,
            File file) {
            this.clazz = clazz;
            this.unmarshallerTweaker = unmarshallerTweaker;
            this.file = file;
        }

        public FileInstanceSupplier(Class<T> clazz, Function<Unmarshaller, Unmarshaller> unmarshallerTweaker,
            InputStream inputStream) {
            this.clazz = clazz;
            this.unmarshallerTweaker = unmarshallerTweaker;
            this.inputStream = inputStream;
        }

        public T get() {
            this.inputStream = this.buildFileStream();
            if (null == this.inputStream) {
                return null;
            } else {
                Stopwatch stopwatch = Stopwatch.createStarted();

                try {
                    JAXBContext context = JaxbContexts.of(this.clazz);
                    Unmarshaller unmarshaller = context.createUnmarshaller();
                    unmarshaller = this.unmarshallerTweaker.apply(unmarshaller);
                    T result = (T) unmarshaller.unmarshal(SaxSources.newSecurityUnmarshalSource(this.inputStream));
                    T var5 = result;
                    return var5;
                } catch (ParserConfigurationException | SAXException | JAXBException var9) {
                    JaxbContexts.LOGGER.error("FAILED to unmarshall file:{}, class:{}", this.file.getName(),
                        this.clazz);
                } finally {
                    Closeables.closeQuietly(this.inputStream);
                    JaxbContexts.LOGGER.debug("unmarshall cost: {}, clazz: {}, file: {}",
                        stopwatch.elapsed(TimeUnit.MILLISECONDS), this.clazz, this.file);
                }

                return null;
            }
        }

        private InputStream buildFileStream() {
            if (null != this.inputStream) {
                return this.inputStream;
            } else {
                if (this.file != null && this.file.exists()) {
                    try {
                        return new FileInputStream(this.file);
                    } catch (FileNotFoundException var2) {
                        JaxbContexts.LOGGER.error("FAILED to unmarshall file:{},class:{}", this.file.getName(),
                            this.clazz, var2);
                    }
                }

                return this.inputStream;
            }
        }
    }
}