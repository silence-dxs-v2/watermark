package com.silence.watermarkdemo.utils;

import cn.hutool.http.HttpRequest;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.util.IOUtils;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * @description：
 * @author：dxs
 * @date：2024/8/13 09:16
 */
@Service
@Slf4j
public class MyTest {
    /**
     * 文档转换器
     */
    @Value("${watermark.host}")
    private String host;
//    @Resource
//    private DocumentConverter documentConverter;
//    public void  word_to_pdf(InputStream stream, OutputStream sourceOutput){
//
//        try  {
//            DocumentFormat type =  DefaultDocumentFormatRegistry.DOCX;
//
//            documentConverter.convert(stream)
//                    .as(type)
//                    .to(sourceOutput)
//                    .as(DefaultDocumentFormatRegistry.PDF)
//                    .execute();
//        } catch (Exception e) {
//            log.error("转换Word文件为PDF时发生错误", e);
//            throw new RuntimeException(e);
//        }
//
//
//    }

    public void wordToPdfByJod(InputStream stream, OutputStream sourceOutput, String encodedFileName) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int bytesRead;


        try {
            while ((bytesRead = stream.read(buffer)) != -1) {
                byteArrayOutputStream.write(buffer, 0, bytesRead);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }


        try {
            InputStream data = HttpRequest.post(host + "/lool/convert-to/pdf").form("data",
                    byteArrayOutputStream.toByteArray(), encodedFileName).execute().bodyStream();


            IOUtils.copy(data, sourceOutput);
            log.info("word转换pdf访问调用成功返回");
        } catch (IOException e) {

            log.error("请求pdf转换访问异常{}", e);
        }


    }
}
