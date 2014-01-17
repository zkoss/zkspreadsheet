package org.zkoss.zss.ng.exporter;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.model.impl.pdf.BackupExporter;
import org.zkoss.zss.model.impl.pdf.PdfExporter;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.ExcelImporter;
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
		
		File temp = Setup.getTempFile("pdfExportTest",".pdf");
		
		exportBook(book, temp);
		
		Util.open(temp);
	}
	
	@Test
	public void exportPdfTest() {
		NBook book = importBook("book/taubman.xlsx");
		int[] colbreak = book.getSheet(0).getViewInfo().getColumnBreaks();
		File temp = Setup.getTempFile("pdfExportTest",".pdf");
		exportBook(book, temp);
		Util.open(temp);
	}
	
	@Test
	public void zss529Test() {
		NBook book = importBook("book/taubman-nobreak.xlsx");
		File temp = Setup.getTempFile("pdfExportTest",".pdf");
		exportBook(book, temp);
		Util.open(temp);
	}
	
	@Test
	public void oldExportPdf() throws IOException {
		BackupExporter exporter = new BackupExporter();
		InputStream is  = PdfExporterTest.class.getResourceAsStream("book/taubman.xlsx");
		ExcelImporter importer = new ExcelImporter();
		XBook book = importer.imports(is, "test");
		is.close();
		File tmpFile = Setup.getTempFile("pdfExportTest",".pdf");
		OutputStream os = new FileOutputStream(tmpFile);
		exporter.export(book, os);
		os.close();
		Util.open(tmpFile);
	}
	
	private NBook importBook(String path) {
		NImporter importer = new ExcelImportFactory().createImporter();
		InputStream is  = PdfExporterTest.class.getResourceAsStream(path);
		NBook book = null;
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
	
	private void exportBook(NBook book, File file) {
		
		PdfExporter exporter = new PdfExporter();
		exporter.enableGridLines(false);
		try {
			exporter.export(book, file);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
