package org.zkoss.zss.model;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBooks;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SFont;
import org.zkoss.zss.model.SRichText;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.SFont.Boldweight;
import org.zkoss.zss.model.SFont.Underline;
import org.zkoss.zss.model.impl.pdf.PdfExporter;
import org.zkoss.zss.range.SImporter;
import org.zkoss.zss.range.impl.imexp.ExcelImportFactory;

public class PdfExporterTest {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}

	@Test
	public void richTextTest() {
		SBook book = SBooks.createBook("rich");
		SSheet sheet = book.createSheet("first");
		SCell cell = sheet.getCell(0, 0);
		
		SRichText rText = cell.setupRichTextValue();
		SFont font1 = book.createFont(true);
		font1.setColor(book.createColor("#0000FF"));
		font1.setStrikeout(true);
		rText.addSegment("abc", font1);
		
		SFont font2 = book.createFont(true);
		font2.setColor(book.createColor("#FF0000"));
		font2.setBoldweight(Boldweight.BOLD);
		rText.addSegment("123", font2);
		
		SFont font3 = book.createFont(true);
		font3.setColor(book.createColor("#C78548"));
		font3.setUnderline(Underline.SINGLE);
		rText.addSegment("xyz", font3);
		
		cell = sheet.getCell(0, 1);
		rText = cell.setupRichTextValue();
		font1 = book.createFont(true);
		font1.setColor(book.createColor("#FFFF00"));
		font1.setItalic(true);
		rText.addSegment("Hello", font1);
		
		font2 = book.createFont(true);
		font2.setColor(book.createColor("#FF33FF"));
		font2.setBoldweight(Boldweight.BOLD);
		rText.addSegment("World", font2);
		
		font3 = book.createFont(true);
		font3.setColor(book.createColor("#CCCC99"));
		font3.setName("HGPSoeiKakupoptai");
		rText.addSegment("000", font3);
		
		File temp = Setup.getTempFile("pdfExportTest",".pdf");
		
		exportBook(book, temp);
		
		Util.open(temp);
	}
	
	@Ignore("manual test only")
	@Test
	public void exportPdfTest() {
		SBook book = importBook("book/PrintSetup.xlsx");
		File temp = Setup.getTempFile("pdfExportTest",".pdf");
		exportBook(book, temp);
		Util.open(temp);
	}
	
	@Ignore("not fixed yet")
	@Test
	public void zss529Test() {
		SBook book = importBook("book/taubman-nobreak.xlsx");
		File temp = Setup.getTempFile("pdfExportTest",".pdf");
		exportBook(book, temp);
		Util.open(temp);
	}
	
	private SBook importBook(String path) {
		SImporter importer = new ExcelImportFactory().createImporter();
		InputStream is  = PdfExporterTest.class.getResourceAsStream(path);
		SBook book = null;
		try {
			book = importer.imports(is, "PDFBook");
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return book;
	}
	
	private void exportBook(SBook book, File file) {
		
		PdfExporter exporter = new PdfExporter();
		try {
			exporter.export(book, file);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
