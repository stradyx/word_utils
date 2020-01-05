package com.strady.word_utils.utils;

import com.aspose.words.Document;
import com.aspose.words.ExportHeadersFootersMode;
import com.aspose.words.HtmlSaveOptions;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

/**
 * aspose-words工具类
 *
 * @Author: strady
 * @Date: 2020-01-03
 * @Time: 14:34:22
 * @Description:
 */
public class AsposeWordUtil {

    public static String parseWord2Html(InputStream stream) throws Exception {
        Document doc = new Document(stream);
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.setExportHeadersFootersMode(ExportHeadersFootersMode.NONE);
        ByteArrayOutputStream htmlStream = new ByteArrayOutputStream();
        String htmlText = "";
        try {
            doc.save(htmlStream, saveOptions);
            htmlText = new String(htmlStream.toByteArray(), "UTF-8");
            htmlStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return htmlText;
    }
}
