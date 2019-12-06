/*
 * Copyright (c) Huawei Technologies Co., Ltd. 2000-2019. All rights reserved.
 */

package com.app.excelcsvconverter.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.Charset;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.OpenOption;
import java.nio.file.Path;
import java.nio.file.StandardOpenOption;
import java.nio.file.attribute.FileAttribute;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.Collection;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SecureFileUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(SecureFileUtil.class);

    private static final String DEFAULT_OSSUSER = "ossuser";

    private static final String NOT_SAVE_PATTERN = "(.*([/\\\\]{1}[\\.\\.]{1,2}|[\\.\\.]{1,2}[/\\\\]{1}|\\.\\.).*|\\.)";

    private static final String PATH_WHITE_LIST
        = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-=[];\\',./ ~!@#$%^&*()_+\"{}|:<>?";

    private SecureFileUtil() {
    }

    public static String normalizeFileName(String targetDir, String fileName) throws IOException {
        File file = new File(targetDir, fileName);
        String canonicalPath = file.getCanonicalPath();
        File targetDirectory = new File(targetDir);
        String canonicalTargetDir = targetDirectory.getCanonicalPath();
        if (canonicalPath.startsWith(canonicalTargetDir)) {
            return canonicalPath;
        } else {
            throw new IllegalStateException("File is outside target directory.");
        }
    }

    public static FileAttribute<Set<PosixFilePermission>> getDefaultFileAttribute(boolean isReadShare) {
        String permissons = isReadShare ? "rw-r-----" : "rw-------";
        Set<PosixFilePermission> perms = PosixFilePermissions.fromString(permissons);
        return PosixFilePermissions.asFileAttribute(perms);
    }

    private static boolean isPosix() {
        return FileSystems.getDefault().supportedFileAttributeViews().contains("posix");
    }

    private static WrapAcl<? extends OpenOption, ?, ?> getLinuxWrapAcl(boolean isRead, boolean isGroupReadShare,
        Path path) {
        LOGGER.info("linux OS");
        Set<OpenOption> options = new HashSet();
        if (!isRead) {
            options.add(StandardOpenOption.CREATE);
            options.add(StandardOpenOption.TRUNCATE_EXISTING);
        }

        FileAttribute<Set<PosixFilePermission>> attr = getDefaultFileAttribute(isGroupReadShare);
        return new WrapAcl(path, options, attr);
    }

    public static OutputStream getFileSafeOutputStream(File file, boolean isGroupReadShare) {
        try {
            if (isPosix()) {
                WrapAcl<? extends OpenOption, ?, ?> wrapAcl = getLinuxWrapAcl(false, isGroupReadShare, file.toPath());
                Path filePath;
                if (!file.exists()) {
                    filePath = Files.createFile(wrapAcl.getPath(), wrapAcl.getAttr());
                } else {
                    filePath = file.toPath();
                }

                return Files.newOutputStream(filePath, (OpenOption[]) wrapAcl.getOptions().toArray(new OpenOption[0]));
            } else {
                return new FileOutputStream(file, false);
            }
        } catch (IOException var4) {
            LOGGER.error("Create outputStream fail.");
        }
        return null;
    }

    public static OutputStream getFileSafeOutputStream(File file) {
        return getFileSafeOutputStream(file, false);
    }

    public static BufferedWriter getFileSafeBufferedWriter(File file, Charset charset, boolean isGroupReadShare) {
        try {
            if (isPosix()) {
                WrapAcl<? extends OpenOption, ?, ?> wrapAcl = getLinuxWrapAcl(false, isGroupReadShare, file.toPath());
                Path path;
                if (!file.exists()) {
                    path = Files.createFile(wrapAcl.getPath(), wrapAcl.getAttr());
                } else {
                    path = file.toPath();
                }

                return Files.newBufferedWriter(path, charset,
                    (OpenOption[]) wrapAcl.getOptions().toArray(new OpenOption[0]));
            } else {
                return Files.newBufferedWriter(file.toPath(), charset);
            }
        } catch (IOException var5) {
            LOGGER.error("Create buffered writer fail.");
        }
        return null;
    }

    public static InputStream getFileSafeInputStream(File file, boolean isGroupReadShare) {
        try {
            if (isPosix()) {
                WrapAcl<? extends OpenOption, ?, ?> wrapAcl = getLinuxWrapAcl(true, isGroupReadShare, file.toPath());
                return Files.newInputStream(wrapAcl.getPath(),
                    (OpenOption[]) wrapAcl.getOptions().toArray(new OpenOption[0]));
            } else {
                return new FileInputStream(file);
            }
        } catch (IOException var3) {
            LOGGER.error("Crate inputstream fail.", var3);
        }
        return null;
    }

    public static OutputStream createFileOutputStream(String fileName, boolean isAppend) throws IOException {
        return createFileOutputStream(new File(fileName), isAppend);
    }

    public static OutputStream createFileOutputStream(File chckeFile) throws IOException {
        return createFileOutputStream(chckeFile, false);
    }

    public static OutputStream createFileOutputStream(File file, boolean isAppend) throws IOException {
        if (isPosix()) {
            Path file1 = Files.createFile(file.toPath(), getDefaultFileAttribute(false));
            Set<OpenOption> options = new HashSet();
            options.add(StandardOpenOption.CREATE);
            if (isAppend) {
                options.add(StandardOpenOption.APPEND);
            }

            return Files.newOutputStream(file1, (OpenOption[]) options.toArray(new OpenOption[0]));
        } else {
            return new FileOutputStream(file, isAppend);
        }
    }

    public static File getSafeFile(String filePath) throws IllegalStateException {
        return getSafeFile(filePath, false);
    }

    public static File getSafeFile(String filePath, boolean replacePathByWhiteList) {
        if (isSafePath(filePath)) {
            String replacedPath = filePath;
            if (replacePathByWhiteList) {
                replacedPath = checkFile(filePath);
            }

            return new File(replacedPath);
        } else {
            throw new RuntimeException("Invalid file path");
        }
    }

    private static String checkFile(String filePath) {
        if (filePath == null) {
            throw new RuntimeException("Invalid file path,path is null");
        } else {
            StringBuilder tmpStrBuf = new StringBuilder();

            for (int i = 0; i < filePath.length(); ++i) {
                for (int j = 0; j
                    < "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-=[];\\',./ ~!@#$%^&*()_+\"{}|:<>?"
                    .length(); ++j) {
                    if (filePath.charAt(i)
                        == "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-=[];\\',./ ~!@#$%^&*()_+\"{}|:<>?"
                        .charAt(j)) {
                        tmpStrBuf.append(
                            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890-=[];\\',./ ~!@#$%^&*()_+\"{}|:<>?"
                                .charAt(j));
                        break;
                    }
                }
            }

            return tmpStrBuf.toString();
        }
    }

    private static boolean isSafePath(String filePath) {
        Pattern pattern = Pattern.compile("(.*([/\\\\]{1}[\\.\\.]{1,2}|[\\.\\.]{1,2}[/\\\\]{1}|\\.\\.).*|\\.)");
        Matcher matcher = pattern.matcher(filePath);
        boolean isSafe = !matcher.matches();
        if (!isSafe) {
            LOGGER.info("isSafePath Invalid path");
        }

        return isSafe;
    }

    public static Path createTempDirectory(Path dir, String prefix) throws IOException {
        return Files.createTempDirectory(dir, prefix);
    }

    private static class WrapAcl<S extends OpenOption, C extends Collection<T>, T> {
        private Path path;

        private Set<S> options;

        private FileAttribute<C> attr;

        WrapAcl(Path path, Set<S> options, FileAttribute<C> attr) {
            this.path = path;
            this.options = options;
            this.attr = attr;
        }

        public Path getPath() {
            return this.path;
        }

        public void setPath(Path path) {
            this.path = path;
        }

        public Set<S> getOptions() {
            return this.options;
        }

        public void setOptions(Set<S> options) {
            this.options = options;
        }

        public FileAttribute<C> getAttr() {
            return this.attr;
        }

        public void setAttr(FileAttribute<C> attr) {
            this.attr = attr;
        }
    }
}
