package com.wanfang.convert.convert;

import com.wanfang.common.convert.file.convert.IHtmlConverter;
import com.wanfang.common.convert.file.tool.PoiTool;

public class test {
    public static void main(String[] args) {
        IHtmlConverter htmlConverter = new PoiTool();
        htmlConverter.docxToHtml("G:\\aa\\aaaaa.docx", "G:\\aa\\aa.html");
    }
}
