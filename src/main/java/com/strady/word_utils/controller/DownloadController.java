package com.strady.word_utils.controller;

import com.deepoove.poi.XWPFTemplate;
import com.strady.word_utils.utils.FreemarkerUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * 使用模板生成word下载controller
 *
 * @Author: strady
 * @Date: 2020-01-04
 * @Time: 10:59:34
 * @Description:
 */
@RestController
@RequestMapping(value = "/download")
public class DownloadController {

    /**
     * 使用poi生成word
     *
     * @param request
     * @param response
     */
    @RequestMapping(value = "/poi")
    public void poi(HttpServletRequest request, HttpServletResponse response) {
        try {
            //输出文件名
            String fname = "work_poi";
            OutputStream os = response.getOutputStream();// 取得输出流
            response.reset();// 清空输出流
            // 对中文文件名的处理
            fname = java.net.URLEncoder.encode(fname, "UTF-8");
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(fname.getBytes("UTF-8"), "GBK") + ".doc");
            response.setContentType("application/msword");// 定义输出类型
            try {//通过map存放要填充的数据
                Map<String, Object> data = new HashMap<String, Object>();
                data.put("name", "张三");//把每项数据写进map，key的命名要与word里面的一样
                data.put("idCard", "123456789123456789");
                data.put("post", "Java工程师");
                data.put("time", new SimpleDateFormat("yyyy年MM月dd日").format(new Date()));

                //调用模板，填充数据
//                InputStream fis = new FileInputStream(new File(request.getSession().getServletContext().getRealPath("/") + "template.docx"));
                InputStream fis = DownloadController.class.getResourceAsStream("/static/template/poi/template.docx");

                XWPFTemplate template = XWPFTemplate.compile(fis).render(data);
                template.write(os);
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 使用freemarker生成word
     *
     * @param request
     * @param response
     */
    @RequestMapping(value = "/freemarker")
    public void freemarker(HttpServletRequest request, HttpServletResponse response) {
        try {

            String fname = "work_freemarker";
            try {
                Map<String, Object> data = new HashMap<String, Object>();//通过map存放要填充的数据
                data.put("name", "张三");//把每项数据写进map，key的命名要与word里面的一样
                data.put("idCard", "123456789123456789");
                data.put("post", "Java工程师");
                data.put("time", new SimpleDateFormat("yyyy年MM月dd日").format(new Date()));
                FreemarkerUtil.createWord(request, response, data, "template.ftl", fname);
            } catch (Exception e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
