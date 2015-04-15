package org.zkoss.zss.issue;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.impl.pdf.PdfExporter;

public class Issue974Test {
	
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

	/**
	 * Test export to pdf with hidden row and column.
	 */
	@Test
	public void exportFillPatterns() {
		Book book = Util.loadBook(this, "book/974-fill-patterns-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue974FillPatternsTest",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		Util.open(temp);
	}
	
	private void exportBook(SBook book, File file) {
		
		PdfExporter pdfExporter = new PdfExporter();
		
		try {
			pdfExporter.export(book, file);
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
