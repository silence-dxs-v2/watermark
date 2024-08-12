package com.silence.watermarkdemo.utils;

import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Element;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfWriter;
import lombok.extern.slf4j.Slf4j;
import org.apache.fontbox.ttf.TrueTypeCollection;
import org.apache.fontbox.ttf.TrueTypeFont;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDFont;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.pdfbox.pdmodel.font.PDType1Font;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.pdmodel.graphics.state.PDExtendedGraphicsState;
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
import javax.swing.*;
import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.charset.StandardCharsets;
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

	public static void putWaterMark(InputStream inputStream, OutputStream outputStream, String content, String fileName) {
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

	public static void putWaterRemarkToExcel(InputStream inputStream, OutputStream outputStream, String content) {
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


//	public static void putWaterRemarkToPDF(InputStream inputStream, OutputStream outputStream, String content) {
//		try {
//			PDDocument document = PDDocument.load(inputStream);
//			Integer width = 600;
//			Integer height = 150;
//			BufferedImage image = getWaterImage(width, height, content, 30);
//			ByteArrayOutputStream os = new ByteArrayOutputStream();
//			ImageIO.write(image, "png", os);
//			PDImageXObject pdImage = PDImageXObject.createFromByteArray(document, os.toByteArray(), "waterMark");
//			for (PDPage page : document.getPages()) {
//				try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true, true)) {
//					contentStream.drawImage(pdImage, 0, 0, width, height);
//				}
//			}
//			document.save(outputStream);
//			document.close();
//		} catch (Exception e) {
//			log.error("putWaterRemarkToPDF error", e);
//		}
//	}

//	public static void putWaterRemarkToPDF(InputStream inputStream, OutputStream outputStream, String content) {
//		try {
//			PDDocument document = PDDocument.load(inputStream);
//			PDExtendedGraphicsState gs = new PDExtendedGraphicsState();
//			gs.setNonStrokingAlphaConstant(0.5f); // 设置透明度
//			gs.setStrokingAlphaConstant(0.5f);
//
//			// 使用PDFBox的默认字体
//			InputStream fontStream = FileRemarkUtil.class.getClassLoader().getResourceAsStream("font/wqy-zenhei.ttc");
//			PDType0Font font= PDType0Font.load(document, fontStream);
//			// 加载提取的字体
////			PDType0Font font = PDType0Font.load(document, fontStream,true);
//
//			for (PDPage page : document.getPages()) {
//				PDRectangle mediaBox = page.getMediaBox();
//				float pageWidth = mediaBox.getWidth();
//				float pageHeight = mediaBox.getHeight();
//
//				try (PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true, true)) {
//					contentStream.setGraphicsStateParameters(gs);
//					contentStream.setNonStrokingColor(Color.GRAY); // 设置水印颜色
//					contentStream.setFont(font, 50);
//
//					// 计算水印位置
//					float textWidth = font.getStringWidth(content) / 1000 * 50;
//					float x = (pageWidth - textWidth) / 2;
//					float y = pageHeight / 2;
//
//					// 添加水印
//					contentStream.beginText();
//					contentStream.setTextMatrix(2, 0, 0, 2, x, y); // 旋转水印
//					contentStream.showText(content);
//					contentStream.endText();
//				}
//			}
//
//			document.save(outputStream);
//			document.close();
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//	}
public static void putWaterRemarkToPDF(InputStream input, OutputStream output, String waterMarkName) {
	BufferedOutputStream bos = null;
	try {
		// 读取文件，生成reader
		com.itextpdf.text.pdf.PdfReader reader = new com.itextpdf.text.pdf.PdfReader(input);
		// 生成输出文件，开启输出流
//		bos = new BufferedOutputStream(new FileOutputStream(new File("/Users/dxs/Desktop/share/new/pdfwatermark")));
		// 读取输出流，生成stamper（印章）
		com.itextpdf.text.pdf.PdfStamper stamper = new com.itextpdf.text.pdf.PdfStamper(reader, output);
		// 设置stamper加密
		stamper.setEncryption(null, "Ka_ze".getBytes(StandardCharsets.UTF_8), PdfWriter.ALLOW_PRINTING, PdfWriter.STANDARD_ENCRYPTION_128);

		// 获取总页数 +1, 下面从1开始遍历
		int total = reader.getNumberOfPages() + 1;
		// 使用classpath下面的字体库
		com.itextpdf.text.pdf.BaseFont base = null;
		try {
			base = com.itextpdf.text.pdf.BaseFont.createFont("STSong-Light", "UniGB-UCS2-H", com.itextpdf.text.pdf.BaseFont.EMBEDDED);
		} catch (Exception e) {
			// 日志处理
			e.printStackTrace();
		}

		// 间隔
		int interval = -15;
		// 获取水印文字的高度和宽度
		int textH = 0, textW = 0;
		JLabel label = new JLabel();
		label.setText(waterMarkName);
		FontMetrics metrics = label.getFontMetrics(label.getFont());
		textH = metrics.getHeight();
		textW = metrics.stringWidth(label.getText());

		// 设置水印透明度
		com.itextpdf.text.pdf.PdfGState gs = new com.itextpdf.text.pdf.PdfGState();
		gs.setFillOpacity(0.2f);
		gs.setStrokeOpacity(0.7f);

		com.itextpdf.text.Rectangle pageSizeWithRotation = null;
		PdfContentByte content = null;
		for (int i = 1; i < total; i++) {
			// 在内容上方加水印
			content = stamper.getOverContent(i);
			// 在内容下方加水印
			// content = stamper.getUnderContent(i);
			content.saveState();
			content.setGState(gs);

			// 设置字体和字体大小
			content.beginText();
			content.setFontAndSize(base, 25);

			// 获取每一页的高度、宽度
			pageSizeWithRotation = reader.getPageSizeWithRotation(i);
			float pageHeight = pageSizeWithRotation.getHeight();
			float pageWidth = pageSizeWithRotation.getWidth();

			// 根据纸张大小多次添加， 水印文字成30度角倾斜
			for (int height = interval + textH; height < pageHeight; height = height + textH * 15) {
				for (int width = interval + textW; width < pageWidth + textW; width = width + textW * 2) {
					content.showTextAligned(Element.ALIGN_LEFT, waterMarkName, width - textW, height - textH, 30);
				}
			}
			content.endText();
		}

		// 关流
		stamper.close();
		reader.close();
	} catch (DocumentException | IOException e) {
		e.printStackTrace();
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
