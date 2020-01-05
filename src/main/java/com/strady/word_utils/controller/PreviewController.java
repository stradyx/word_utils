package com.strady.word_utils.controller;

import com.strady.word_utils.utils.AsposeWordUtil;
import com.strady.word_utils.utils.PoiUtil;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.nio.charset.Charset;

/**
 * word预览controller
 *
 * @Author: strady
 * @Date: 2020-01-04
 * @Time: 10:59:34
 * @Description:
 */
@RestController
@RequestMapping(value = "/preview")
public class PreviewController {

    /**
     * 使用poi预览word
     *
     * @param request
     * @param httpServletResponse
     * @return
     */
    @RequestMapping(value = "/poi")
    public ResponseEntity<String> poi(HttpServletRequest request, HttpServletResponse httpServletResponse) {
        HttpHeaders responseHeaders = new HttpHeaders();
        responseHeaders.setContentType(new MediaType("text", "html", Charset.forName("UTF-8")));
        try {
            MultipartHttpServletRequest mulRequest = (MultipartHttpServletRequest) request;
            //获取上传文件
            MultipartFile file = mulRequest.getFile("file");
            //获取文件名
            String filename = file.getOriginalFilename();
            if (null == filename || "".equals(filename)) {
                return new ResponseEntity<String>("请选择要上传的文件", responseHeaders, HttpStatus.OK);
            }
            if (!filename.endsWith(".doc") && !filename.endsWith(".docx")) {
                return new ResponseEntity<String>("上传文件的格式不正确", responseHeaders, HttpStatus.OK);
            }
            //将文件转换为inputStream
            InputStream input = file.getInputStream();
            String result = "word文件解析失败";
            //判断文件格式选择不同解析方法
            if (!filename.endsWith(".doc")) {
                result = PoiUtil.word2007ToHtml(file);
            } else {
                result = PoiUtil.word2003ToHtml(input);
            }
            // TODO 打印解析后html字符串
            return new ResponseEntity<String>(result, responseHeaders, HttpStatus.OK);
        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<String>("word文件解析失败", responseHeaders, HttpStatus.OK);
        }
    }

    /**
     * 使用aspone-words预览word
     *
     * @param request
     * @param httpServletResponse
     * @return
     */
    @RequestMapping(value = "/aspose")
    public ResponseEntity<String> aspose(HttpServletRequest request, HttpServletResponse httpServletResponse) {
        HttpHeaders responseHeaders = new HttpHeaders();
        responseHeaders.setContentType(new MediaType("text", "html", Charset.forName("UTF-8")));
        try {
            MultipartHttpServletRequest mulRequest = (MultipartHttpServletRequest) request;
            MultipartFile file = mulRequest.getFile("file");
            String filename = file.getOriginalFilename();

            if (null == filename || "".equals(filename)) {
                return new ResponseEntity<String>("请选择要上传的文件", responseHeaders, HttpStatus.OK);
            }
            if (!filename.endsWith(".doc") && !filename.endsWith(".docx")) {
                return new ResponseEntity<String>("上传文件的格式不正确", responseHeaders, HttpStatus.OK);
            }

            InputStream input = file.getInputStream();
            String result = "word文件解析失败";
            result = AsposeWordUtil.parseWord2Html(input);
            // TODO 打印解析后html字符串
            return new ResponseEntity<String>(result, responseHeaders, HttpStatus.OK);
        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<String>("word文件解析失败", responseHeaders, HttpStatus.OK);
        }
    }

}
