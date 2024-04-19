package com.kanseiu.office.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.Comparator;
import java.util.Enumeration;
import java.util.stream.Stream;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

public class CustomizeFileUtil {

    /**
     * 复制文件
     * @param sourceFilePath        待复制的文件
     * @param destinationDirPath    复制目的目录
     */
    public static void copyFile(String sourceFilePath, String destinationDirPath) {
        Path sourcePath = Paths.get(sourceFilePath);
        Path destinationPath = Paths.get(destinationDirPath);

        // 检查源文件是否存在
        if (!Files.exists(sourcePath)) {
            System.err.println("Source file does not exist, skipping copy.");
            return;
        }

        try {
            // 如果目的文件夹不存在，则递归创建
            if (Files.notExists(destinationPath)) {
                Files.createDirectories(destinationPath);
            }

            // 构建目标文件的完整路径
            Path targetPath = destinationPath.resolve(sourcePath.getFileName());

            // 执行复制
            Files.copy(sourcePath, targetPath, StandardCopyOption.REPLACE_EXISTING);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 将 压缩文件（xlsx）解压为 文件夹
     * @param zipFilePath   压缩文件路径
     * @param destDir       输出文件夹路径
     */
    public static void unzip(String zipFilePath, String destDir) {
        File dir = new File(destDir);
        // 创建输出目录如果它不存在
        if (!dir.exists()) {
            if(!dir.mkdirs()) {
                throw new RuntimeException("创建文件夹失败！");
            }
        }

        // 缓冲区大小
        byte[] buffer = new byte[1024];
        try (ZipFile zipFile = new ZipFile(zipFilePath)){
            Enumeration<? extends ZipEntry> zipEntries = zipFile.entries();
            while (zipEntries.hasMoreElements()) {
                ZipEntry entry = zipEntries.nextElement();
                File newFile = newFile(dir, entry);
                // 创建所有不存在的父目录
                if (entry.isDirectory()) {
                    if (!newFile.isDirectory() && !newFile.mkdirs()) {
                        throw new IOException("Failed to create directory " + newFile);
                    }
                } else {
                    // 创建所有不存在的父目录
                    File parent = newFile.getParentFile();
                    if (!parent.isDirectory() && !parent.mkdirs()) {
                        throw new IOException("Failed to create directory " + parent);
                    }

                    // 写文件内容
                    try (FileOutputStream fos = new FileOutputStream(newFile)){
                        InputStream is = zipFile.getInputStream(entry);
                        int len;
                        while ((len = is.read(buffer)) > 0) {
                            fos.write(buffer, 0, len);
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取文件夹，并压缩为xlsx文件
     * @param srcDir        数据源文件夹
     * @param xlsxFilePath  输出文件
     */
    public static void zipFolderToXlsx(String srcDir, String xlsxFilePath){
        Path sourceDirPath = Paths.get(srcDir);
        try (ZipOutputStream zos = new ZipOutputStream(new FileOutputStream(xlsxFilePath))) {
            Files.walkFileTree(sourceDirPath, new SimpleFileVisitor<>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    zos.putNextEntry(new ZipEntry(sourceDirPath.relativize(file).toString()));
                    Files.copy(file, zos);
                    zos.closeEntry();
                    return FileVisitResult.CONTINUE;
                }

                @Override
                public FileVisitResult preVisitDirectory(Path dir, BasicFileAttributes attrs) throws IOException {
                    if (!sourceDirPath.equals(dir)) {
                        zos.putNextEntry(new ZipEntry(sourceDirPath.relativize(dir) + "/"));
                        zos.closeEntry();
                    }
                    return FileVisitResult.CONTINUE;
                }
            });
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 删除文件/文件夹及其子文件
     */
    public static void deleteFileOrDir(String filePath){
        Path directoryToBeDeleted = Paths.get(filePath);
        if(Files.exists(directoryToBeDeleted)) {
            try (Stream<Path> streamPath = Files.walk(directoryToBeDeleted)){
                streamPath.sorted(Comparator.reverseOrder()).map(Path::toFile).forEach(File::delete);
                if (Files.notExists(directoryToBeDeleted)) {
                    System.out.println("文件或文件夹删除成功！");
                } else {
                    throw new RuntimeException("文件或文件夹删除失败！");
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.out.println("文件或文件不存在，不需要删除！");
        }
    }

    /**
     * 创建文件
     */
    private static File newFile(File destinationDir, ZipEntry zipEntry) throws IOException {
        File destFile = new File(destinationDir, zipEntry.getName());
        String destDirPath = destinationDir.getCanonicalPath();
        String destFilePath = destFile.getCanonicalPath();
        if (!destFilePath.startsWith(destDirPath + File.separator)) {
            throw new IOException("Entry is outside of the target directory: " + zipEntry.getName());
        }
        return destFile;
    }
}
