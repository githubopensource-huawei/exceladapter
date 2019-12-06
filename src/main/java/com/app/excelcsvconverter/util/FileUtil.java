/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.util;

import com.google.common.io.Files;
import com.google.common.io.LineProcessor;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.file.CopyOption;
import java.nio.file.FileVisitResult;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.StandardCopyOption;
import java.nio.file.attribute.BasicFileAttributes;
import java.nio.file.attribute.FileAttribute;

public final class FileUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(FileUtil.class);

    private static final String CROSS_PATH1 = "..\\";

    private static final String CROSS_PATH2 = "../";

    private FileUtil() {
    }

    public static String securityFileName(String name) {
        String tmp = name;

        int lastLen;
        do {
            lastLen = tmp.length();
            tmp = tmp.replace("..\\", "").replace("../", "");
        } while (tmp.length() != lastLen);

        return tmp;
    }

    public static String getCanonicalPath(File file) {
        String path = "";

        try {
            path = file.getCanonicalPath();
        } catch (FileNotFoundException var3) {
            LOGGER.error("{} not found", file.getName());
        } catch (IOException var4) {
            LOGGER.error(var4.getMessage(), var4);
        }

        return path;
    }

    public static String getExtension(File file) {
        return Files.getFileExtension(file.getName());
    }

    public static String getBaseName(File file) {
        String fileName = file.getName();
        int lastDotIndex = fileName.lastIndexOf(46);
        return lastDotIndex == -1 ? fileName : fileName.substring(0, lastDotIndex);
    }

    public static <T> T readLines(File file, Charset charset, LineProcessor<T> processor) throws IOException {
        BufferedReader reader = java.nio.file.Files.newBufferedReader(file.toPath(), charset);
        Throwable var4 = null;

        T var6;
        try {
            var6 = processor.getResult();
        } catch (Throwable var15) {
            var4 = var15;
            throw var15;
        } finally {
            if (reader != null) {
                if (var4 != null) {
                    try {
                        reader.close();
                    } catch (Throwable var14) {
                        var4.addSuppressed(var14);
                    }
                } else {
                    reader.close();
                }
            }

        }

        return var6;
    }

    public static void copyFileToDirectory(File srcFile, File destDir) throws IOException {
        copyFile(srcFile, new File(destDir, srcFile.getName()));
    }

    public static void copyFileToDirectory(File srcFile, File destDir, String newName) throws IOException {
        copyFile(srcFile, new File(destDir, newName));
    }

    public static void copyFile(File srcFile, File destFile) throws IOException {
        File destDir = destFile.getParentFile();
        boolean succeed = destDir.exists() || destDir.mkdirs();
        if (!succeed) {
            LOGGER.error("FAILED to mkdirs: ", destDir);
        } else if (srcFile.equals(destFile)) {
            LOGGER.warn("source file same as destination file, NO need to copy: {}", srcFile.getName());
        } else {
            java.nio.file.Files.copy(srcFile.toPath(), destFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
        }
    }

    public static void copyFileRecursive(File fromPath, File toPath) throws IOException {
        if (!toPath.exists() && !toPath.mkdir()) {
            LOGGER.info("makeDir fail :" + getCanonicalPath(toPath));
        } else {
            File[] files = fromPath.listFiles();
            if (files != null) {
                File[] var3 = files;
                int var4 = files.length;

                for (int var5 = 0; var5 < var4; ++var5) {
                    File subfile = var3[var5];
                    if (subfile.isFile()) {
                        Files.copy(subfile, new File(toPath, subfile.getName()));
                    } else if (subfile.isDirectory()) {
                        copyFileRecursive(subfile, new File(toPath, subfile.getName()));
                    }
                }

            }
        }
    }

    public static void copyDirectoryToDirectory(File srcDir, File destDir) throws IOException {
        boolean succeed = destDir.exists() || destDir.mkdirs();
        if (!succeed) {
            LOGGER.error("FAILED to mkdirs: {}", destDir.getName());
        } else {
            copyDirectoryContent(srcDir, new File(destDir, srcDir.getName()));
        }
    }

    public static void copyDirectoryContent(File srcDir, File destDir) throws IOException {
        boolean succeed = destDir.exists() || destDir.mkdirs();
        if (!succeed) {
            LOGGER.error("FAILED to mkdirs: {}", destDir.getName());
        } else {
            copy(srcDir.toPath(), destDir.toPath());
        }
    }

    public static void copy(Path from, Path to) throws IOException {
        java.nio.file.Files.walkFileTree(from, new CopyingFileVisitor(from, to));
    }

    public static boolean isWritable(String path) {
        if (path != null && !path.isEmpty()) {
            File f = new File(path);
            if (!f.exists()) {
                return true;
            } else {
                try {
                    if (!f.canWrite() && !f.setWritable(true)) {
                        return false;
                    }
                } catch (Exception var17) {
                    LOGGER.error("isOpenByOthers: ", var17);
                    return false;
                }

                FileOutputStream os = null;

                boolean var4;
                try {
                    os = new FileOutputStream(path, true);
                    return true;
                } catch (FileNotFoundException var15) {
                    LOGGER.error("File is not found!");
                    var4 = false;
                } finally {
                    if (os != null) {
                        try {
                            os.close();
                        } catch (IOException var14) {
                            LOGGER.error("isOpenByOthers: ", var14);
                        }
                    }

                }

                return var4;
            }
        } else {
            return false;
        }
    }

    public static boolean deleteFile(File file) {
        boolean before = null != file && file.isFile();
        deleteFileQuietly(file);
        boolean after = null != file && !file.isFile();
        return before && after;
    }

    public static void deleteFileQuietly(File file) {
        if (null != file && file.isFile()) {
            deleteQuietly(file.toPath());
        }

    }

    public static boolean deleteDirectory(File dir) {
        boolean before = null != dir && dir.isDirectory();
        deleteDirectoryQuietly(dir);
        boolean after = null != dir && !dir.isDirectory();
        return before && after;
    }

    public static void deleteDirectoryQuietly(File dir) {
        if (null != dir && dir.isDirectory()) {
            deleteQuietly(dir.toPath());
        }

    }

    public static void deleteQuietly(Path path) {
        if (null != path) {
            try {
                java.nio.file.Files.walkFileTree(path, new DeletingFileVisitor());
            } catch (FileNotFoundException var2) {
                LOGGER.error("{} not found", path.getFileName());
            } catch (IOException var3) {
                LOGGER.error("FAILED to delete quietly: {}", path.getFileName(), var3);
            }

        }
    }

    private static class DeletingFileVisitor extends SimpleFileVisitor<Path> {
        private DeletingFileVisitor() {
        }

        public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
            if (attrs.isRegularFile()) {
                LOGGER.trace("deleting regular file: {}", file);
                java.nio.file.Files.delete(file);
            }

            return FileVisitResult.CONTINUE;
        }

        public FileVisitResult visitFileFailed(Path file, IOException exc) {
            LOGGER.error("FAILED to visit path: " + file, exc);
            return FileVisitResult.CONTINUE;
        }

        public FileVisitResult postVisitDirectory(Path dir, IOException exc) throws IOException {
            LOGGER.trace("deleting directory: {}", dir);
            java.nio.file.Files.delete(dir);
            return FileVisitResult.CONTINUE;
        }
    }

    private static class CopyingFileVisitor extends SimpleFileVisitor<Path> {
        private Path from;

        private Path to;

        private StandardCopyOption copyOption;

        public CopyingFileVisitor(Path from, Path to) {
            this(from, to, StandardCopyOption.REPLACE_EXISTING);
        }

        public CopyingFileVisitor(Path from, Path to, StandardCopyOption copyOption) {
            this.from = from;
            this.to = to;
            this.copyOption = copyOption;
        }

        public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
            Path target = this.to.resolve(this.from.relativize(dir));
            if (!java.nio.file.Files.exists(target)) {
                java.nio.file.Files.createDirectory(target);
            }

            return FileVisitResult.CONTINUE;
        }

        public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
            java.nio.file.Files.copy(file, this.to.resolve(this.from.relativize(file)), this.copyOption);
            return FileVisitResult.CONTINUE;
        }
    }
}