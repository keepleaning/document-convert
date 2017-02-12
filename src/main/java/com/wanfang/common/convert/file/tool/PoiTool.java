package com.wanfang.common.convert.file.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.util.Arrays;
import java.util.List;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.converter.core.BasicURIResolver;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.w3c.dom.Document;

import com.wanfang.common.convert.core.exception.NotImplementedException;
import com.wanfang.common.convert.core.exception.XyException;
import com.wanfang.common.convert.file.convert.IHtmlConverter;

/**
 * 使用POI作为转换器实现转换功能
 */
public class PoiTool implements IHtmlConverter {
    private List<String> supportedExtensionsForHtmlConverter;
    /**
     * 存放图片的路径
     */
    private final static String IMAGE_PATH = "image";

    public PoiTool() {
        supportedExtensionsForHtmlConverter = Arrays.asList("doc", "docx", "xls");
    }

    @Override
    public void docToHtml(String sourceFileName, String targetFileName) {
        initFolder(targetFileName);
        String imagePathStr = initImageFolder(targetFileName);
        try {
            HWPFDocument wordDocument = new HWPFDocument(new FileInputStream(sourceFileName));
            Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
            WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(document);
            HtmlPicturesManager picturesManager = new HtmlPicturesManager(imagePathStr, IMAGE_PATH);
            wordToHtmlConverter.setPicturesManager(picturesManager);
            wordToHtmlConverter.processDocument(wordDocument);
            Document htmlDocument = wordToHtmlConverter.getDocument();
            DOMSource domSource = new DOMSource(htmlDocument);
            StreamResult streamResult = new StreamResult(new File(targetFileName));

            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
        } catch (Exception e) {
            throw new XyException("将doc文件转换为html时出错!" + e.getMessage());
        }
    }

    @Override
    public void docxToHtml(String sourceFileName, String targetFileName) {
        initFolder(targetFileName);
        String imagePathStr = initImageFolder(targetFileName);
        OutputStreamWriter outputStreamWriter = null;
        try {
            XWPFDocument document = new XWPFDocument(new FileInputStream(sourceFileName));
            XHTMLOptions options = XHTMLOptions.create();
            // 存放图片的文件夹
            options.setExtractor(new FileImageExtractor(new File(imagePathStr)));
            // html中图片的路径
            options.URIResolver(new BasicURIResolver(IMAGE_PATH));
            outputStreamWriter = new OutputStreamWriter(new FileOutputStream(targetFileName), "utf-8");
            XHTMLConverter xhtmlConverter = (XHTMLConverter) XHTMLConverter.getInstance();
            xhtmlConverter.convert(document, outputStreamWriter, options);
        } catch (Exception e) {
            throw new XyException("将docx文件转换为html时出错!" + e.getMessage());
        } finally {
            try {
                if (outputStreamWriter != null) {
                    outputStreamWriter.close();
                }
            } catch (Exception e) {
            }
        }
    }

    @Override
    public void xlsToHtml(String sourceFileName, String targetFileName) {
        initFolder(targetFileName);
        try {
            Document doc = ExcelToHtmlConverter.process(new File(sourceFileName));
            DOMSource domSource = new DOMSource(doc);
            StreamResult streamResult = new StreamResult(new File(targetFileName));
            TransformerFactory tf = TransformerFactory.newInstance();
            Transformer serializer = tf.newTransformer();
            serializer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
            serializer.setOutputProperty(OutputKeys.INDENT, "yes");
            serializer.setOutputProperty(OutputKeys.METHOD, "html");
            serializer.transform(domSource, streamResult);
        } catch (Exception e) {
            throw new XyException("将xls文件转换为html时出错!" + e.getMessage());
        }
    }

    @Override
    public void xlsxToHtml(String sourceFileName, String targetFileName) {
        throw new NotImplementedException();
    }

    /**
     * 初始化存放html文件的文件夹
     * 
     * @param targetFileName
     *            html文件的文件名
     */
    private void initFolder(String targetFileName) {
        File targetFile = new File(targetFileName);
        if (targetFile.exists()) {
            targetFile.delete();
        }
        String targetPathStr = targetFileName.substring(0, targetFileName.lastIndexOf(File.separator));
        File targetPath = new File(targetPathStr);
        // 如果文件夹不存在，则创建
        if (!targetPath.exists()) {
            targetPath.mkdirs();
        }
    }

    /**
     * 初始化存放图片的文件夹
     * 
     * @param htmlFileName
     *            html文件的文件名
     * @return 存放图片的文件夹路径
     */
    private String initImageFolder(String htmlFileName) {
        String targetPathStr = htmlFileName.substring(0, htmlFileName.lastIndexOf(File.separator));
        // 创建存放图片的文件夹
        String imagePathStr = targetPathStr + File.separator + IMAGE_PATH + File.separator;
        File imagePath = new File(imagePathStr);
        if (imagePath.exists()) {
            imagePath.delete();
        }
        imagePath.mkdir();
        return imagePathStr;
    }

    @Override
    public List<String> getSupportedExtensionsForHtml() {
        return supportedExtensionsForHtmlConverter;
    }

    @Override
    public void setSupportedExtensionsForHtml(List<String> supportedExtensionsForHtmlConverter) {
        this.supportedExtensionsForHtmlConverter = supportedExtensionsForHtmlConverter;
    }

    /**
     * 读取doc文件的文本，不带格式
     * 
     * @param fileName
     *            文件名
     * @return 文件的文本内容
     */
    public static String readTextForDoc(String fileName) {
        String text;
        try (FileInputStream in = new FileInputStream(fileName); WordExtractor wordExtractor = new WordExtractor(in);) {
            text = wordExtractor.getText();
        } catch (Exception e) {
            throw new XyException("读取文件出错", e);
        }
        return text;
    }

    /**
     * 读取docx文件的内容，只读取纯文本内容，不带格式
     * 
     * @param fileName
     *            文件名
     * @return 文件内容
     */
    public static String readTextForDocx(String fileName) {
        String text = null;
        POIXMLTextExtractor ex = null;
        try {
            OPCPackage oPCPackage = POIXMLDocument.openPackage(fileName);
            XWPFDocument xwpf = new XWPFDocument(oPCPackage);
            ex = new XWPFWordExtractor(xwpf);
            text = ex.getText();
        } catch (Exception e) {
            throw new XyException("读取文件出错", e);
        } finally {
            try {
                if (ex != null) {
                    ex.close();
                }
            } catch (Exception e) {
                throw new XyException("关闭流出错:" + fileName, e);
            }
        }
        return text;
    }

}