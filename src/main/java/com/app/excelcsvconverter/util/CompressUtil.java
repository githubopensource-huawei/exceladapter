/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.util;

import com.google.common.base.Strings;

import org.apache.commons.compress.archivers.ArchiveException;
import org.apache.commons.compress.archivers.ArchiveInputStream;
import org.apache.commons.compress.archivers.ArchiveStreamFactory;
import org.apache.commons.compress.archivers.tar.TarArchiveEntry;
import org.apache.commons.compress.archivers.tar.TarArchiveInputStream;
import org.apache.commons.compress.archivers.tar.TarArchiveOutputStream;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Enumeration;
import java.util.zip.GZIPInputStream;
import java.util.zip.GZIPOutputStream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

/**
 * 功能描述：
 *
 * @since 2019-08-19
 */
public class CompressUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(CompressUtil.class);

    private static final int BUFFER_SIZE = 2 * 1024;

    public static String uniqueDirectoryName(String moduleName) {
        StringBuilder nameBuilder = new StringBuilder(moduleName);
        if (!Strings.isNullOrEmpty(moduleName)) {
            nameBuilder.append("_");
        }
        return nameBuilder.append(datetimeText()).toString();
    }

    private static String datetimeText() {
        return (new SimpleDateFormat("yyyy_MM_dd_HH_mm_ss")).format(new Date());
    }

    public static String unZipFiles(File zipFile, String descDir) throws IOException {
        // 解决中文文件夹乱码
        ZipFile zip;
        try {
            zip = new ZipFile(zipFile, EncodingUtil.getDefaultEncoding());
        } catch (ZipException e) {
            LOGGER.error("Unzip file is fail.");
            throw new ZipException("Unzip file is fail.");
        }

        String name = zip.getName().substring(zip.getName().lastIndexOf('\\') + 1, zip.getName().lastIndexOf('.'));
        File pathFile = new File(descDir + File.separator + name);
        if (!pathFile.exists()) {
            pathFile.mkdirs();
        }

        for (Enumeration<? extends ZipEntry> entries = zip.entries(); entries.hasMoreElements(); ) {
            ZipEntry entry = entries.nextElement();
            String zipEntryName = entry.getName();
            InputStream in = zip.getInputStream(entry);
            String outPath = (descDir + "/" + name + "/" + zipEntryName).replaceAll("\\*", "/");

            // 判断路径是否存在,不存在则创建文件路径
            File file = new File(outPath.substring(0, outPath.lastIndexOf('/')));
            if (!file.exists()) {
                file.mkdirs();
            }
            // 判断文件全路径是否为文件夹,如果是上面已经上传,不需要解压
            if (new File(outPath).isDirectory()) {
                continue;
            }

            FileOutputStream out = new FileOutputStream(outPath);
            byte[] buf1 = new byte[1024];
            int len;
            while ((len = in.read(buf1)) > 0) {
                out.write(buf1, 0, len);
            }
            in.close();
            out.close();
        }
        return pathFile.getAbsolutePath();
    }

    public static void zipFiles(String srcDir, File file, boolean KeepDirStructure) throws Exception {

        try (FileOutputStream fileOutputStream = new FileOutputStream(file); ZipOutputStream zos = new ZipOutputStream(fileOutputStream)) {
            File sourceFile = new File(srcDir);
            compress(sourceFile, zos, "", KeepDirStructure);
        } catch (Exception e) {
            throw new Exception("zip error from ZipUtils", e);
        }
    }

    private static void compress(File sourceFile, ZipOutputStream zos, String name, boolean KeepDirStructure) throws Exception {
        byte[] buf = new byte[BUFFER_SIZE];
        if (sourceFile.isFile()) {
            zos.putNextEntry(new ZipEntry(name));
            int len;
            FileInputStream in = new FileInputStream(sourceFile);
            while ((len = in.read(buf)) != -1) {
                zos.write(buf, 0, len);
            }
            zos.closeEntry();
            in.close();
        } else {
            File[] listFiles = sourceFile.listFiles();
            if (listFiles == null || listFiles.length == 0) {
                if (KeepDirStructure) {
                    zos.putNextEntry(new ZipEntry(name));
                    zos.closeEntry();
                }
            } else {
                for (File file : listFiles) {
                    if (KeepDirStructure) {
                        String fileName;
                        if (!file.getPath().contains(".")) {
                            fileName = name + file.getName() + "/";
                        } else {
                            fileName = name + file.getName();
                        }
                        compress(file, zos, fileName, KeepDirStructure);
                    } else {
                        compress(file, zos, file.getName(), KeepDirStructure);
                    }
                }
            }
        }
    }

    public static void tarFiles(File srcFile, File destFile) throws Exception {

        TarArchiveOutputStream taos = new TarArchiveOutputStream(new FileOutputStream(destFile));
        taos.setLongFileMode(TarArchiveOutputStream.LONGFILE_GNU);
        File[] files = srcFile.listFiles();
        for (File file : files) {
            archive(file, taos, "");
        }
        taos.flush();
        taos.close();
    }

    private static void archive(File srcFile, TarArchiveOutputStream taos, String basePath) throws Exception {

        if (srcFile.isDirectory()) {
            archiveDir(srcFile, taos, basePath);
        } else {
            archiveFile(srcFile, taos, basePath);
        }
    }

    private static void archiveFile(File file, TarArchiveOutputStream taos, String dir) throws Exception {

        TarArchiveEntry entry = new TarArchiveEntry(dir + file.getName());

        entry.setSize(file.length());

        taos.putArchiveEntry(entry);

        BufferedInputStream bis = new BufferedInputStream(new FileInputStream(file));
        int count;
        byte data[] = new byte[BUFFER_SIZE];
        while ((count = bis.read(data, 0, BUFFER_SIZE)) != -1) {
            taos.write(data, 0, count);
        }

        bis.close();

        taos.closeArchiveEntry();
    }

    private static void archiveDir(File dir, TarArchiveOutputStream taos, String basePath) throws Exception {

        File[] files = dir.listFiles();

        if (files.length < 1) {
            TarArchiveEntry entry = new TarArchiveEntry(basePath + dir.getName() + File.separator);

            taos.putArchiveEntry(entry);
            taos.closeArchiveEntry();
        }

        for (File file : files) {

            // 递归归档
            archive(file, taos, basePath + dir.getName() + File.separator);

        }
    }

    public static File unTarFiles(File tarFile, File dstFolder) throws Exception {
        long maxSize = 500 * 1024 * 1024;
        String name = tarFile.getName().substring(tarFile.getName().lastIndexOf('\\') + 1, tarFile.getName().lastIndexOf('.'));
        File pathFile = new File(dstFolder + File.separator + name);
        if (!pathFile.exists()) {
            pathFile.mkdirs();
        }
        try (FileInputStream fis = new FileInputStream(tarFile); TarArchiveInputStream tarIn = new TarArchiveInputStream(fis)) {
            TarArchiveEntry entry;
            long total = 0;
            while ((entry = tarIn.getNextTarEntry()) != null) {
                File file = new File(pathFile, getSecurityName(entry));
                // 为解决空目录
                if (entry.isDirectory()) {
                    if (!file.mkdirs()) {
                        LOGGER.error("mkdirs failed for {}", file);
                    }
                    continue;
                }
                File parentDir = file.getParentFile();
                if (!parentDir.exists() && !parentDir.mkdirs()) {
                    LOGGER.error("mkdirs failed for {}", parentDir);
                    continue;
                }

                try (OutputStream fos = SecureFileUtil.getFileSafeOutputStream(file); BufferedOutputStream bos = new BufferedOutputStream(fos)) {
                    int count;
                    byte[] buffer = new byte[4096];
                    while (total <= maxSize && (count = tarIn.read(buffer, 0, buffer.length)) != -1) {
                        bos.write(buffer, 0, count);
                        total = total + count;
                    }
                    if (total > maxSize) {
                        LOGGER.error("The size of unzip file is over {}", maxSize);
                        throw new Exception();
                    }
                    bos.flush();
                } catch (Exception e1) {
                    LOGGER.error("unTar fail", e1);
                    throw new Exception(e1);
                }
            }
        } catch (FileNotFoundException e) {
            LOGGER.error("unTar fail, File is not Found");
            throw new Exception(e);
        } catch (IOException e) {
            LOGGER.error("unTar fail", e);
            throw new Exception(e);
        }
        return pathFile.getCanonicalFile();
    }

    private static String getSecurityName(TarArchiveEntry entry) {
        return FileUtil.securityFileName(entry.getName());
    }

    /**
     * 解压缩多个文件压缩的.TAR.GZ文件
     *
     * @return
     */
    public static File unGzFiles(File gzFile, File dstFolder) throws Exception {
        if (!dstFolder.isDirectory() && !dstFolder.mkdirs()) {
            LOGGER.error("mkdirs failde for {}", dstFolder);
            throw new Exception();
        }
        String fileName = gzFile.getName().substring(0, gzFile.getName().length() - 7);
        File pathFile = new File(dstFolder + File.separator + fileName);
        if (!pathFile.exists()) {
            pathFile.mkdirs();
        }
        try (FileInputStream fis = new FileInputStream(gzFile);
            GZIPInputStream gis = new GZIPInputStream(new BufferedInputStream(fis));
            ArchiveInputStream in = new ArchiveStreamFactory().createArchiveInputStream(ArchiveStreamFactory.TAR, gis);
            BufferedInputStream bufferedInputStream = new BufferedInputStream(in)) {
            TarArchiveEntry entry = (TarArchiveEntry) in.getNextEntry();
            while (entry != null) {
                String name = getSecurityName(entry);
                String[] names = name.split("/");
                File file = pathFile;
                for (String str : names) {
                    file = new File(file, Strings.nullToEmpty(str));
                }
                if (name.endsWith("/")) {
                    if (!file.mkdirs()) {
                        LOGGER.error("mkdirs failed for {}", file);
                        continue;
                    }
                } else {
                    writeFile(file, bufferedInputStream);
                }
                entry = (TarArchiveEntry) in.getNextEntry();
            }
        } catch (ArchiveException e) {
            LOGGER.error("fail to uncompress ", e);
            throw new Exception(e);
        } catch (FileNotFoundException e) {
            LOGGER.error("fail to uncompress for file is not found");
            throw new Exception(e);
        } catch (IOException e) {
            LOGGER.error("fail to uncompress", e);
            throw new Exception(e);
        }
        return pathFile;
    }

    public static void gzTar(File tarFile) throws Exception {
        try (FileInputStream fileInputStream = new FileInputStream(tarFile);
            BufferedInputStream bufferedInput = new BufferedInputStream(fileInputStream);
            FileOutputStream fileOutputStream = new FileOutputStream(tarFile + ".gz");
            GZIPOutputStream gzipOutputStream = new GZIPOutputStream(fileOutputStream);
            BufferedOutputStream bufferedOutput = new BufferedOutputStream(gzipOutputStream)) {
            byte[] cache = new byte[1024];
            for (int index = bufferedInput.read(cache); index != -1; index = bufferedInput.read(cache)) {
                bufferedOutput.write(cache, 0, index);
            }
        } catch (Exception e) {
            throw new Exception(e);
        } finally {
            FileUtil.deleteFileQuietly(tarFile);
        }
    }

    private static void writeFile(File file, BufferedInputStream bufferedInputStream) throws Exception {
        File parentFile = file.getParentFile();
        if (!parentFile.exists()) {
            if (!parentFile.mkdirs()) {
                LOGGER.error("mkdirs failed for {}", parentFile);
                return;
            }
        }
        try (OutputStream outputFile = SecureFileUtil.getFileSafeOutputStream(file); BufferedOutputStream bufferedOutputStream = new BufferedOutputStream(outputFile)) {
            byte[] buf = new byte[4096];
            int cnt;
            while ((cnt = bufferedInputStream.read(buf)) != -1) {
                bufferedOutputStream.write(buf, 0, cnt);
            }
            bufferedOutputStream.flush();
        } catch (Exception e) {
            LOGGER.error("writeFile error:", e);
            throw new Exception(e);
        }
    }

}
















