package com.strady.word_utils.utils;

import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @Author: strady
 * @Date: 2020-01-04
 * @Time: 11:11:53
 * @Description:
 */
public class PoiUtil {
    /**
     * 将word 2007转换为html代码
     *
     * @param file
     * @return
     * @throws IOException
     */
    public static String word2007ToHtml(MultipartFile file) throws IOException {
        try {
            if (file.isEmpty() || file.getSize() <= 0) {
                return "解析的word文件不能为空";
            } else {
                // 加载word文档生成 XWPFDocument对象
                InputStream in = file.getInputStream();
                XWPFDocument document = new XWPFDocument(in);
                // 使用字符数组流获取解析的内容
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                XHTMLConverter.getInstance().convert(document, baos, null);
                String content = baos.toString();
                baos.close();
                byte[] lens = baos.toByteArray();
                content = new String(lens);
                return content;
            }
        } catch (Exception e) {
            return "word文件解析失败";
        }
    }

    /**
     * 将word 2003转换为html代码
     *
     * @param file
     * @return
     * @throws IOException
     */
    public static String word2003ToHtml(InputStream file) throws IOException {
        HWPFDocument wordDocument = new HWPFDocument((InputStream) file);
        WordToHtmlConverter wordToHtmlConverter = null;
        try {
            wordToHtmlConverter = new WordToHtmlConverter(
                    DocumentBuilderFactory.newInstance().newDocumentBuilder()
                            .newDocument());
            wordToHtmlConverter.processDocument(wordDocument);
            Document htmlDocument = wordToHtmlConverter.getDocument();
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(out);

            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
            out.close();
            byte[] lens = out.toByteArray();
            String content = new String(lens, "utf-8");
            return content;
        } catch (Exception e) {
            e.printStackTrace();
            return "word文件解析失败";
        }
    }
}