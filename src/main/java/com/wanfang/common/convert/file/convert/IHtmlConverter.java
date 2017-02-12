package com.wanfang.common.convert.file.convert;

import java.util.List;

import com.wanfang.common.convert.core.exception.XyException;
import com.wanfang.common.convert.file.tool.FileTool;

/**
 * 将其它格式的文件转换为html
 */
public interface IHtmlConverter {

	/**
	 * 文件转换为html文件
	 * 
	 * @param sourceFileName
	 *            源文件的文件名
	 * @param targetFileName
	 *            目标文件的文件名
	 */
	default void toHtml(String sourceFileName, String targetFileName) {
		String extension = FileTool.getFileExtension(sourceFileName);
		switch (extension) {
		case "doc":
			docToHtml(sourceFileName, targetFileName);
			break;
		case "docx":
			docxToHtml(sourceFileName, targetFileName);
			break;
		case "xls":
			xlsToHtml(sourceFileName, targetFileName);
			break;
		case "xlsx":
			xlsxToHtml(sourceFileName, targetFileName);
			break;
		default:
			throw new XyException("不支持的文件类型:" + extension);
		}
	}

	/**
	 * doc文件转换为html文件
	 * 
	 * @param sourceFileName
	 *            源文件的文件名
	 * @param targetFileName
	 *            目标文件的文件名
	 */
	void docToHtml(String sourceFileName, String targetFileName);

	/**
	 * docx文件转换为html文件
	 * 
	 * @param sourceFileName
	 *            源文件的文件名
	 * @param targetFileName
	 *            目标文件的文件名
	 */
	void docxToHtml(String sourceFileName, String targetFileName);

	/**
	 * 
	 * @return 获得Html转换器支持的文件后缀名,不区分大小写，""表示没有扩展名的文件
	 */
	List<String> getSupportedExtensionsForHtml();

	/**
	 * 判断转换器是否支持一个扩展名
	 * 
	 * @param extension
	 *            文件扩展名
	 * @return true 支持，false 不支持
	 */
	default boolean isSupportedExtensionsForHtml(String extension) {
		List<String> list = getSupportedExtensionsForHtml();
		return list != null && list.stream().anyMatch(item -> item.equalsIgnoreCase(extension));
	}

	/**
	 * 
	 * @param extensions
	 *            设置Html转换器支持的文件后缀名,不区分大小写，""表示没有扩展名的文件
	 */
	void setSupportedExtensionsForHtml(List<String> extensions);

	/**
	 * xls文件转换为html文件
	 * 
	 * @param sourceFileName
	 *            源文件的文件名
	 * @param targetFileName
	 *            目标文件的文件名
	 */
	void xlsToHtml(String sourceFileName, String targetFileName);

	/**
	 * xlsx文件转换为html文件
	 * 
	 * @param sourceFileName
	 *            源文件的文件名
	 * @param targetFileName
	 *            目标文件的文件名
	 */
	void xlsxToHtml(String sourceFileName, String targetFileName);

}
