package org.zkoss.zss.ng.exporter;

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
import org.zkoss.zss.model.impl.html.HtmlExporter;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngapi.impl.imexp.ExcelImportFactory;
import org.zkoss.zss.ngmodel.NBook;

public class HtmlExporterTest {
	
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
	
	@Ignore("manual test only")
	@Test
	public void exportHtml() {
		NBook book = importBook("book/simpleChart.xlsx");
		File temp = Setup.getTempFile("htmlExportTest",".html");
		exportBook(book, temp);
		Util.open(temp);
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
		
		HtmlExporter exporter = new HtmlExporter();
		try {
			exporter.export(book, file);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
