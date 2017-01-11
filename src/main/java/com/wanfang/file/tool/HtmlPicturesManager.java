package com.wanfang.file.tool;

import java.io.FileOutputStream;

import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.usermodel.PictureType;

public class HtmlPicturesManager implements PicturesManager {
	private String imageRealPath;
	private String imageHtmlPath;

	/**
	 * @param imageRealPath
	 *            存放图片的文件夹名称
	 * @param imageHtmlPath
	 *            图片在html文件中的路径
	 */
	public HtmlPicturesManager(String imageRealPath, String imageHtmlPath) {
		this.imageHtmlPath = imageHtmlPath;
		this.imageRealPath = imageRealPath;
	}

	@Override
	public String savePicture(byte[] content, PictureType pictureType, String name, float widthInches,
			float heightInches) {
		try (FileOutputStream out = new FileOutputStream(imageRealPath + name)) {
			out.write(content);
		} catch (Exception e) {
		}
		return imageHtmlPath + "/" + name;
	}

}
