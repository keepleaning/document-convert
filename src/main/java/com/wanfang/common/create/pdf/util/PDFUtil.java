package com.wanfang.common.create.pdf.util;

import java.io.BufferedWriter;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.StringWriter;
import java.nio.charset.Charset;
import java.util.Map;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.Element;
import com.itextpdf.text.ExceptionConverter;
import com.itextpdf.text.FontFactory;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.tool.xml.XMLWorkerHelper;
import com.wanfang.common.create.word.util.WORDUtil;

import freemarker.template.Configuration;
import freemarker.template.Template;

/**
 * PDF生成工具类
 * 
 * @author zhangsh 创建时间：2016年1月25日 下午1:19:27
 */
public class PDFUtil {
    // 支持中文
    public static final String fontChinese = "STSong-Light";
    public static final String encodingChinese = "UniGB-UCS2-H";

    public static String generate(String template, Map<String, Object> variables, String templatePackage)
            throws Exception {
        Configuration config = new Configuration();
        config.setDefaultEncoding("utf-8");
        config.setClassForTemplateLoading(WORDUtil.class, templatePackage);
        Template tp = config.getTemplate(template);
        StringWriter stringWriter = new StringWriter();
        BufferedWriter writer = new BufferedWriter(stringWriter);
        tp.setEncoding("UTF-8");
        tp.process(variables, writer);
        String htmlStr = stringWriter.toString();
        writer.flush();
        writer.close();
        return htmlStr;
    }

    /**
     * 生成pdf
     * 
     * @param file
     * @throws Exception
     */
    public String createPdf(Map<String, Object> variables, String url, String templateName, String templatePackge)
            throws Exception {
        File file = new File(url.substring(0, url.lastIndexOf("/")));
        // 如果文件夹不存在则创建
        if (!file.exists() && !file.isDirectory()) {
            file.mkdirs();
        }
        // 创建Document对象
        Document document = new Document(PageSize.A4, 50, 50, 50, 50);
        // 创建书写对象
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(url));
        writer.setPageEvent(new PageXofY());
        // 打开文档
        document.open();
        String htmlStr = generate(templateName, variables, templatePackge);
        InputStream is = new ByteArrayInputStream(htmlStr.getBytes("UTF-8"));
        XMLWorkerHelper.getInstance().parseXHtml(writer, document, is, Charset.forName("UTF-8"));
        document.close();
        return url;
    }

    // 生成底部页码
    public class PageXofY extends PdfPageEventHelper {
        /** 这个PdfTemplate实例用于保存总页数 */
        protected PdfTemplate total;
        /** 页码字体 */
        protected BaseFont helv;

        @Override
        public void onOpenDocument(PdfWriter writer, Document document) {
            total = writer.getDirectContent().createTemplate(100, 100);
            total.setBoundingBox(new Rectangle(-20, -20, 100, 100));
            try {
                helv = BaseFont.createFont(fontChinese, encodingChinese, BaseFont.NOT_EMBEDDED);
            } catch (Exception e) {
                throw new ExceptionConverter(e);
            }
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            PdfContentByte cb = writer.getDirectContent();
            cb.saveState();
            String text = "第 " + writer.getPageNumber() + " 页  共";
            float textBase = document.bottom() - 20;
            float textSize = helv.getWidthPoint(text, 10);
            cb.beginText();
            cb.setFontAndSize(helv, 10);

            float adjust = helv.getWidthPoint("0", 10);
            cb.setTextMatrix((document.right() - textSize + adjust * 6) / 2, textBase);
            cb.showText(text);
            cb.endText();
            cb.addTemplate(total, (document.right() + textSize + adjust * 6) / 2 + adjust, textBase);

            cb.restoreState();
        }

        @Override
        public void onCloseDocument(PdfWriter writer, Document document) {
            total.beginText();
            total.setFontAndSize(helv, 10);
            total.setTextMatrix(0, 0);
            total.showText(String.valueOf(writer.getPageNumber()) + " 页");
            total.endText();
        }
    }

    // 生成水印
    public class Watermark extends PdfPageEventHelper {

        public Phrase watermark = new Phrase("",
                FontFactory.getFont(fontChinese, encodingChinese, 58, 0, BaseColor.LIGHT_GRAY));

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            PdfContentByte canvas = writer.getDirectContent();
            ColumnText.showTextAligned(canvas, Element.ALIGN_CENTER, watermark, 298, 421, 45);
        }
    }

}