package com.silence.watermarkdemo.utils;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;


import lombok.extern.slf4j.Slf4j;


import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.List;

@Slf4j
public class Documents4jUtil {

    public static void test_word_to_pdf(InputStream stream,OutputStream sourceOutput) {


    }

    /**
     * word转pdf
     *
     */
    public static void convertWordToPdf(InputStream stream,OutputStream sourceOutput) {
        String os = System.getProperty("os.name").toLowerCase();
        log.info("convertWordToPdf 当前操作系统：{}", os);
        if (os.contains("win")) {
            // Windows操作系统
            windowsWordToPdf(stream,sourceOutput);
        } else if (os.contains("nix") || os.contains("nux") || os.contains("mac")) {
            // Unix/Linux/Mac操作系统
            linuxWordToPdf(stream,sourceOutput);
//            try {
//                wordToPdf(stream,sourceOutput);
//            } catch (Exception e) {
//                e.printStackTrace();
//            }
        } else {
            // 未知操作系统
            throw new RuntimeException("不支持当前操作系统转换文档。");
        }
    }

    /**
     * 通过documents4j 实现word转pdf -- Windows 环境 需要有 Microsoft Office 服务
     *
     */
    public static void windowsWordToPdf(InputStream stream, OutputStream sourceOutput) {
        try{
            IConverter converter = LocalConverter.builder().build();
            converter.convert(stream)
                    .as(DocumentType.DOCX)
                    .to(sourceOutput)
                    .as(DocumentType.PDF).execute();
        } catch (Exception e) {
            log.error("winWordToPdf windows环境word转换为pdf时出现异常：", e);
        }
    }

    /**
     * 通过libreoffice 实现word转pdf -- linux 环境 需要有 libreoffice 服务
     *
     */


    public static void linuxWordToPdf(InputStream stream,OutputStream sourceOutput) {
        // 创建临时文件
        File tempFile = createTempFileFromInputStream(stream);
        // 构建LibreOffice的命令行工具命令
        String command = "soffice --headless --invisible --convert-to pdf " + tempFile.getAbsolutePath() + " --outdir " + tempFile.getParent();
        // 执行转换命令
        try {
            if (!executeLinuxCmd(command)) {
                throw new IOException("转换失败");
            }
            readPdfFileToByteArrayOutputStream(tempFile,sourceOutput);
        } catch (Exception e) {
            log.error("ConvertWordToPdf: Linux环境word转换为pdf时出现异常："+e +tempFile.getPath());
            // 清理临时文件
            tempFile.delete();
        } finally {
            File pdfFile = new File(tempFile.getParent(), tempFile.getName().replace(".docx", ".pdf"));
            //清理转换后的pdf文件
            pdfFile.delete();
            // 清理临时文件，无论是否成功转换
            tempFile.delete();
        }
    }




    /**
     * 执行命令行
     *
     * @param cmd 命令行
     * @return
     * @throws IOException
     */
    private static boolean executeLinuxCmd(String cmd) throws IOException {
        Process process = Runtime.getRuntime().exec(cmd);
        try {
            process.waitFor();
        } catch (InterruptedException e) {
            log.error("executeLinuxCmd 执行Linux命令异常：", e);
            Thread.currentThread().interrupt();
            return false;
        }
        return true;
    }

    /**
     *
     * 创建临时文件
     */
    private static File createTempFileFromInputStream(InputStream inputStream) {
        try {
            File tempFile = File.createTempFile("temp_word", ".docx");
            Files.copy(inputStream, tempFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
            return tempFile;
        } catch (IOException e) {
            log.error("创建临时文件失败：", e);
            throw new RuntimeException("创建临时文件失败", e);
        }
    }

    /**
     * 读取pdf文件
     */
    private static void readPdfFileToByteArrayOutputStream(File tempFile,OutputStream sourceOutput){
        try {
            Path outputFile = Paths.get(tempFile.getParent(), tempFile.getName().replace(".docx", ".pdf"));
            File file = outputFile.toFile();
            FileInputStream fileInputStream = new FileInputStream(file);
//            Files.copy(outputFile, sourceOutput);
//            FileRemarkUtil.putWaterMark(fileInputStream, sourceOutput, "物联网2024-08-12 12:23:23", file.getName());
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
