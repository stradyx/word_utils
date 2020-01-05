package com.strady.word_utils.utils;

import freemarker.template.Configuration;
import freemarker.template.Template;
import sun.misc.BASE64Encoder;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Map;

/**
 * Freemarker生成Word工具类
 *
 * @Author: strady
 * @Date: 2019-12-30
 * @Time: 17:03:47
 * @Description:
 */
public class FreemarkerUtil {

    /**
     * 使用模板，生成word，直接下载
     *
     * @param request
     * @param response
     * @param data         word中需要展示的动态数据，用map集合来保存
     * @param templateName word模板名称，例如：template.ftl
     * @param fileName     生成的文件名称
     */
    @SuppressWarnings("unchecked")
    public static void createWord(HttpServletRequest request,
                                  HttpServletResponse response,
                                  Map data,
                                  String templateName,
                                  String fileName) {
        try {
            //创建配置实例
            Configuration configuration = new Configuration();
            //设置编码
            configuration.setDefaultEncoding("UTF-8");
            //ftl模板文件路径
            configuration.setDirectoryForTemplateLoading(new File(FreemarkerUtil.class.getResource("/static/template/freemarker/").getFile()));
            //configuration.setDirectoryForTemplateLoading(new File(request.getSession().getServletContext().getRealPath("/") + "static/template/freemarker/"));
            //获取模板
            Template template = configuration.getTemplate(templateName);
            //输出流
            fileName = java.net.URLEncoder.encode(fileName, "UTF-8");
            // 清空输出流
            //response.reset();
            response.setHeader("Content-Disposition",
                    "attachment;filename=" + new String(fileName.getBytes("UTF-8"), "GBK") + ".doc");
            response.setContentType("application/msword");// 定义输出类型
            // 将模板和数据模型合并生成文件
            Writer out = new BufferedWriter(new OutputStreamWriter(response.getOutputStream()));
            // 生成文件
            template.process(data, out);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 将图片转换为BASE64为字符串
     *
     * @param filename
     * @return
     * @throws IOException
     */
    public static String getImageString(String filename) throws IOException {
        InputStream in = null;
        byte[] data = null;
        try {
            in = new FileInputStream(filename);
            data = new byte[in.available()];
            in.read(data);
            in.close();
        } catch (IOException e) {
            throw e;
        } finally {
            if (in != null) in.close();
        }
        BASE64Encoder encoder = new BASE64Encoder();
        return data != null ? encoder.encode(data) : "";
    }

}
