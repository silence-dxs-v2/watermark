package com.silence.watermarkdemo.controller;

import com.silence.watermarkdemo.utils.Documents4jUtil;
import com.silence.watermarkdemo.utils.FileRemarkUtil;
import com.silence.watermarkdemo.utils.MyTest;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.thymeleaf.util.DateUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.net.URLEncoder;

/**
 * @description：
 * @author：dxs
 * @date：2024/8/9 10:16
 */
@RestController
@RequestMapping("/watermark")
public class TestController {
//    @Autowired
//    private MyTest myTest;
    @Autowired
    private FileRemarkUtil fileRemarkUtil;
    @PostMapping("/down")
    public void downloadFile(HttpServletResponse response, MultipartFile file){
        // 将文件流拷贝到输出流
        InputStream fileInputStream = null;
        try{
            fileInputStream = file.getInputStream();
            response.setContentType("application/octet-stream");
            String encodedFileName = URLEncoder.encode(file.getOriginalFilename(), "UTF-8")
                    .replaceAll("\\+", "%20");

            response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + encodedFileName);

            fileRemarkUtil.putWaterMark(fileInputStream,response.getOutputStream(), "silence", file.getOriginalFilename());
            response.flushBuffer();
            //更新文件下载次数

        }catch (Exception e){
            e.printStackTrace();

        }finally {
            try {
                fileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    @PostMapping("/wordToPdf")
    public void wordToPdf(HttpServletResponse response, MultipartFile file){
        // 将文件流拷贝到输出流
        InputStream fileInputStream = null;
        try{
            fileInputStream = file.getInputStream();
            response.setContentType("application/octet-stream");
            String encodedFileName = URLEncoder.encode(file.getOriginalFilename(), "UTF-8")
                    .replaceAll("\\+", "%20");

            response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + encodedFileName);

//            Documents4jUtil.convertWordToPdf(fileInputStream,response.getOutputStream());
//            Documents4jUtil.test_word_to_pdf(fileInputStream,response.getOutputStream());
//            myTest.wordToPdfByJod(fileInputStream,response.getOutputStream(),encodedFileName);
            response.flushBuffer();
            //更新文件下载次数

        }catch (Exception e){
            e.printStackTrace();

        }finally {
            try {
                fileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }



}
