package org.zkoss.zss.ng.exporter;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.model.impl.pdf.PdfExporter;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.imexp.ExcelImportFactory;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBooks;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NFont.Boldweight;
import org.zkoss.zss.ngmodel.NFont.Underline;
import org.zkoss.zss.ngmodel.NRichText;
import org.zkoss.zss.ngmodel.NSheet;

public class PdfExporterTest {
	
	private static String DEFAULT_EXPORT_PATH = "./target/test.pdf";
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssContextLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
	}

	@Test
	public void richTextTest() {
		NBook book = NBooks.createBook("rich");
		NSheet sheet = book.createSheet("first");
		NCell cell = sheet.getCell(0, 0);
		
		NRichText rText = cell.setupRichText();
		NFont font1 = book.createFont(true);
		font1.setColor(book.createColor("#0000FF"));
		font1.setStrikeout(true);
		rText.addSegment("abc", font1);
		
		NFont font2 = book.createFont(true);
		font2.setColor(book.createColor("#FF0000"));
		font2.setBoldweight(Boldweight.BOLD);
		rText.addSegment("123", font2);
		
		NFont font3 = book.createFont(true);
		font3.setColor(book.createColor("#C78548"));
		font3.setUnderline(Underline.SINGLE);
		rText.addSegment("xyz", font3);
		
		cell = sheet.getCell(0, 1);
		rText = cell.setupRichText();
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

		exportBook(book);
	}
	
	@Test
	public void exportPdf() {
		NBook book = importBook("./book/simple.xlsx");
		exportBook(book);
	}
	
	private NBook importBook(String path) {
		NImporter importer = new ExcelImportFactory().createImporter();
		InputStream is  = PdfExporterTest.class.getResourceAsStream("./book/simple.xlsx");
		NBook book = null;
		try {
			book = importer.imports(is, "PDFBook");
			is.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return book;
	}
	
	private void exportBook(NBook book) {
		exportBook(book, null);
	}
	
	private void exportBook(NBook book, String path) {
		
		path = path != null ? path : DEFAULT_EXPORT_PATH;
		
		PdfExporter exporter = new PdfExporter();
		try {
			exporter.export(book, new File(path));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
