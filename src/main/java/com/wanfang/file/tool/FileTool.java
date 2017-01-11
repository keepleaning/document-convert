package com.wanfang.file.tool;

public class FileTool {
	private FileTool() {
	}

	/**
	 * 从文件名中截取文件扩展名
	 * 
	 * @param fileName
	 *            文件名
	 * @return 文件扩展名的小写形式，如果文件没有扩展名，则返回长度为0的空字符串
	 */
	public static String getFileExtension(String fileName) {
		int index = fileName.lastIndexOf('.');
		String extension = "";
		if (index > 0) {
			extension = fileName.substring(index + 1).toLowerCase();
		}
		return extension;
	}

}
