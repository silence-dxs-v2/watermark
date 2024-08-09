package com.silence.watermarkdemo.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.thymeleaf.util.DateUtils;

import javax.imageio.ImageIO;
import javax.servlet.ServletOutputStream;
import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.Iterator;
import java.util.Properties;

@Slf4j
public class FileRemarkUtil {
	/**
	 * Excel加水印
	 *
	 * @param filePath 文件路径
	 * @param content  内容
	 */
	public static void putWaterRemarkToExcel(String filePath, String content) {
		try {
			File file = new File(filePath);
			FileInputStream in = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(in);
			FileOutputStream out = new FileOutputStream(filePath);

			Iterator<Sheet> sheetIterator = workbook.sheetIterator();

			Integer width = 800;
			Integer height = 200;

			BufferedImage image = getWaterImage(width, height, content, 45);
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			ImageIO.write(image, "png", os);

			int pictureIdx = workbook.addPicture(os.toByteArray(), Workbook.PICTURE_TYPE_PNG);
			POIXMLDocumentPart poixmlDocumentPart = workbook.getAllPictures().get(pictureIdx);

			while (sheetIterator.hasNext()) {
//				XSSFSheet sheet = (XSSFSheet) sheetIterator.next();
//				String rID = sheet.addRelation(null, XSSFRelation.IMAGES, workbook.getAllPictures().get(pictureIdx)).getRelationship().getId();
//				//设置背景图片水印
//				sheet.getCTWorksheet().addNewPicture().setId(rID);
				XSSFSheet sheet = (XSSFSheet) sheetIterator.next();

				PackagePartName ppn = poixmlDocumentPart.getPackagePart().getPartName();
				String relType = XSSFRelation.IMAGES.getRelation();
				//add relation from sheet to the picture data
				PackageRelationship pr = sheet.getPackagePart().addRelationship(ppn, TargetMode.INTERNAL, relType, null);
				//set background picture to sheet
				sheet.getCTWorksheet().addNewPicture().setId(pr.getId());

			}
			workbook.write(out);

		} catch (Exception e) {
			log.error("putWaterRemarkToExcel error", e);
		}
	}

	/**
	 * 生成水印图片
	 *
	 * @param width   宽度
	 * @param height  高度
	 * @param content 水印内容
	 * @return {@link BufferedImage}
	 */
	public static BufferedImage getWaterImage(Integer width, Integer height, String content, int fontsize) {
		Integer fontStyle = Font.PLAIN;
		Integer fontSize = fontsize;
		try(InputStream fontStream = FileRemarkUtil.class.getClassLoader().getResourceAsStream("font/wqy-zenhei.ttc")) {
			if (fontStream == null) {
				throw new RuntimeException("Font file not found in resources.");
			}
			// 注册本地字体
			Font font1 = Font.createFont(Font.TRUETYPE_FONT, fontStream);
			GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
			ge.registerFont(font1);
			// 设置类型
			Font font = font1.deriveFont(fontStyle, fontSize);
			String fontName = font.getFontName();
			String context_code = new String("物联网".getBytes(), FileRemarkUtil.getSystemFileCharset());
			context_code = context_code+content;

			BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);
			log.info("选择的字体是:{}", fontName);





//			font = new Font("WenQuanYi Zen Hei", fontStyle, fontSize);
			Graphics2D g2d = image.createGraphics(); // 获取Graphics2d对象
			image = g2d.getDeviceConfiguration().createCompatibleImage(width, height, Transparency.TRANSLUCENT);
			g2d.dispose();
			g2d = image.createGraphics();
			//设置字体颜色和透明度
			g2d.setColor(new Color(80, 80, 80, 80));
			// 设置字体
			g2d.setStroke(new BasicStroke(1));
			// 设置字体类型  加粗 大小
			g2d.setFont(font);
			g2d.rotate(Math.toRadians(-10), (double) image.getWidth() / 2, (double) image.getHeight() / 2);//设置倾斜度
			FontRenderContext context = g2d.getFontRenderContext();
			Rectangle2D bounds = font.getStringBounds(context_code, context);
			double x = (width - bounds.getWidth()) / 2;
			double y = (height - bounds.getHeight()) / 2;
			double ascent = -bounds.getY();
			double baseY = y + ascent;
			// 写入水印文字原定高度过小，所以累计写水印，增加高度
			g2d.drawString(context_code, (int) x, (int) baseY);
			// 设置透明度
			g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER));
			// 释放对象
			g2d.dispose();
			return image;
		} catch (Exception e) {
			log.error("getWaterImage error", e);

		}
		return null;

	}

	public static String getSystemFileCharset(){
		Properties pro = System.getProperties();
		log.info("当前系统编码：{}",pro.getProperty("file.encoding"));
		return pro.getProperty("file.encoding");
	}

	public static void putWaterMark(InputStream inputStream, ServletOutputStream outputStream, String content, String fileName) {
		try {
			String fileType = getFileType(fileName);
			if (fileType.equalsIgnoreCase("xlsx")) {
				putWaterRemarkToExcel(inputStream, outputStream, content);
			} else if (fileType.equalsIgnoreCase("docx")) {
//				putWaterRemarkToWord(inputStream, outputStream, content);
				WatermarkUtil.waterMarkDocXDocument(inputStream, outputStream, content);
			} else if (fileType.equalsIgnoreCase("pdf")) {
				putWaterRemarkToPDF(inputStream, outputStream, content);
			} else {
				IOUtils.copy(inputStream, outputStream);
			}
		} catch (Exception e) {
			log.error("putWaterMark error", e);
		}
	}

	public static void putWaterRemarkToExcel(InputStream inputStream, ServletOutputStream outputStream, String content) {
		try {

			XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
			Iterator<Sheet> sheetIterator = workbook.sheetIterator();

			Integer width = 800;
			Integer height = 200;
			BufferedImage image = getWaterImage(width, height, content, 45);
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			ImageIO.write(image, "png", os);
			int pictureIdx = workbook.addPicture(os.toByteArray(), Workbook.PICTURE_TYPE_PNG);
			while (sheetIterator.hasNext()) {
				XSSFSheet sheet = (XSSFSheet) sheetIterator.next();
				String rID = sheet.addRelation(null, XSSFRelation.IMAGES, workbook.getAllPictures().get(pictureIdx)).getRelationship().getId();
				sheet.getCTWorksheet().addNewPicture().setId(rID);
			}
			workbook.write(outputStream);
		} catch (Exception e) {
			log.error("putWaterRemarkToExcel error", e);
		}
	}


	public static void putWaterRemarkToPDF(InputStream inputStream, ServletOutputStream outputStream, String content) {
		try {
			PDDocument document = PDDocument.load(inputStream);
			Integer width = 600;
			Integer height = 150;
			BufferedImage image = getWaterImage(width, height, content, 30);
			ByteArrayOutputStream os = new ByteArrayOutputStream();
			ImageIO.write(image, "png", os);
			PDImageXObject pdImage = PDImageXObject.createFromByteArray(document, os.toByteArray(), "waterMark");
			for (PDPage page : document.getPages()) {
				try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true, true)) {
					contentStream.drawImage(pdImage, 0, 0, width, height);
				}
			}
			document.save(outputStream);
			document.close();
		} catch (Exception e) {
			log.error("putWaterRemarkToPDF error", e);
		}
	}

	private static String getFileType(String filePath) {
		String fileType = "";
		if (filePath.lastIndexOf(".") > -1) {
			fileType = filePath.substring(filePath.lastIndexOf(".") + 1);
		}
		return fileType;
	}

	public static void main(String[] args) {
		putWaterRemarkToExcel("C:\\Users\\Bonnie\\Downloads\\专网号段导入模板.xlsx", "水印 " + DateUtils.createNow().toString());
	}
}
