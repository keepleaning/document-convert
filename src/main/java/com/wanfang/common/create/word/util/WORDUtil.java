package com.wanfang.common.create.word.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.HashMap;
import java.util.Map;
import java.util.UUID;

import freemarker.template.Configuration;
import freemarker.template.Template;
import sun.misc.BASE64Encoder;

/**
 * 类名称：HtmlGenerator
 * 
 * @author zhangsh 创建时间：2016年1月25日 下午1:19:27
 */
public class WORDUtil {
    private static Configuration configuration = null;
    private static Map<String, Template> allTemplates = null;
    public static String type = "report";

    private WORDUtil() {
        throw new AssertionError();
    }

    public static String createDoc(Map<?, ?> dataMap, String url, String templateName, String templatePackage) {
        configuration = new Configuration();
        configuration.setDefaultEncoding("utf-8");
        configuration.setClassForTemplateLoading(WORDUtil.class, templatePackage);
        allTemplates = new HashMap<>(); // Java 7 钻石语法
        try {
            allTemplates.put(type, configuration.getTemplate(templateName));
        } catch (IOException e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
        File f = new File(url);
        File file = new File(url.substring(0, url.lastIndexOf("/")));
        file.setWritable(true);
        // 如果文件夹不存在则创建
        if (!file.exists() && !file.isDirectory()) {
            file.mkdirs();
        }
        Template t = allTemplates.get(type);
        try {
            // 这个地方不能使用FileWriter因为需要指定编码类型否则生成的Word文档会因为有无法识别的编码而无法打开
            Writer w = new OutputStreamWriter(new FileOutputStream(f), "utf-8");
            t.process(dataMap, w);
            w.close();
        } catch (Exception ex) {
            ex.printStackTrace();
            throw new RuntimeException(ex);
        }
        return url;
    }

    public static String getImgLabel(String base64String) {
        String resultDataPicWord = "";
        if ("1".equals(base64String)) {
            resultDataPicWord = "<w:rPr><w:rFonts w:ascii=\"华文仿宋\" w:h-ansi=\"华文仿宋\" w:fareast=\"华文仿宋\" w:cs=\"Arial\" w:hint=\"default\"/><w:kern w:val=\"0\"/></w:rPr><w:t>暂无数据</w:t>";
        } else {
            String uuID = UUID.randomUUID().toString();
            String picWordStart = "<w:pict><w:binData w:name=\"wordml://" + uuID + ".jpg\">";
            String picWordEnd = "</w:binData><v:shape id=\"pic1\" o:spid=\"_x0000_s1026\" o:spt=\"75\" alt=\"\" type=\"#_x0000_t75\" style=\"height:262.5pt;width:450pt;\" filled=\"f\" "
                    + "o:preferrelative=\"t\" stroked=\"f\" coordsize=\"21600,21600\"><v:path/><v:fill on=\"f\" focussize=\"0,0\"/><v:stroke on=\"f\" joinstyle=\"miter\"/><v:imagedata src=\"wordml://"
                    + uuID + ".jpg\" o:title=\"\"/>"
                    + "<o:lock v:ext=\"edit\" aspectratio=\"t\"/><w10:wrap type=\"none\"/><w10:anchorlock/></v:shape></w:pict>";
            resultDataPicWord = picWordStart + base64String + picWordEnd;
        }
        return resultDataPicWord;
    }

    // 将图片转换成BASE64字符串
    @SuppressWarnings("restriction")
    public static String getImageString(String filename) throws IOException {
        InputStream in = null;
        byte[] data = null;
        try {
            File newFile = new File(filename);
            if (newFile.exists()) {
                in = new FileInputStream(newFile);
            }
            // in = new FileInputStream(filename);
            if (in != null) {
                data = new byte[in.available()];
                in.read(data);
                in.close();
            }

        } catch (IOException e) {
            throw e;
        } finally {
            if (in != null)
                in.close();
        }
        BASE64Encoder encoder = new BASE64Encoder();
        return data != null ? encoder.encode(data) : "";
    }

    public static void word2Html() {

    }

    public static void main(String[] args) {
        try {
            System.out.println(getImageString("G:/report/img/1.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}