package com.silence.watermarkdemo.controller;

import com.silence.watermarkdemo.utils.FileRemarkUtil;
import com.silence.watermarkdemo.utils.Word2PdfUtil;
import io.swagger.annotations.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
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
@Api(value = "水印接口", tags = {"水印接口"})
public class TestController {
//    @Autowired
//    private MyTest myTest;
    @Autowired
    private FileRemarkUtil fileRemarkUtil;
    @PostMapping("/down")
    @ApiOperation(value = "word/excel/pdf加水印接口", notes = "word/excel/pdf加水印接口")
//    @ApiImplicitParams({
//            @ApiImplicitParam(name = "file", value = "待处理的文件", required = true, dataType = "org.springframework.web.multipart.MultipartFile", paramType = "form")
//    })
    public void downloadFile(HttpServletResponse response,  @RequestPart(name = "file")
            MultipartFile file){
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
    @ResponseBody
    @ApiOperation(value = "word转pdf接口", notes = "word转pdf接口")
    public ResponseEntity<?> wordToPdf(HttpServletResponse response, @RequestPart(name = "file")MultipartFile file){
        // 将文件流拷贝到输出流
        InputStream fileInputStream = null;
        int i = 0;
        try{
            fileInputStream = file.getInputStream();
//            ServletOutputStream outputStream = response.getOutputStream();
//            if(i==0){
//            response.setContentType("application/octet-stream");
            String encodedFileName = URLEncoder.encode(file.getOriginalFilename(), "UTF-8")
                    .replaceAll("\\+", "%20");
//
//            response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''" + encodedFileName);

//            Documents4jUtil.convertWordToPdf(fileInputStream,response.getOutputStream());
//            Documents4jUtil.test_word_to_pdf(fileInputStream,response.getOutputStream());
//            myTest.wordToPdfByJod(fileInputStream,response.getOutputStream(),encodedFileName);
                ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
                Word2PdfUtil.doc2pdf(fileInputStream, byteArrayOutputStream);


//            response.flushBuffer();
                if(encodedFileName.contains(".docx")){
                    encodedFileName = encodedFileName.replace(".docx", ".pdf");
                }
                if (encodedFileName.contains(".doc")){
                    encodedFileName = encodedFileName.replace(".doc", ".pdf");
                }
                return ResponseEntity.ok()
                        .contentType(MediaType.APPLICATION_OCTET_STREAM)
                        .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename*=UTF-8''" + encodedFileName )
                        .body(byteArrayOutputStream.toByteArray());
//            }

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

        return ResponseEntity.internalServerError().contentType(MediaType.APPLICATION_JSON).body("fail");
    }



}
