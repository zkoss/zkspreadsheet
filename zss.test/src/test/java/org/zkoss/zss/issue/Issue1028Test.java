package org.zkoss.zss.issue;

import java.io.File;
import java.io.IOException;
import java.util.Locale;

import junit.framework.Assert;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.lang.Library;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.impl.pdf.PdfExporter;

public class Issue1028Test {
	
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
	 * Test render long text that across the cell boundary
	 */
	@Test
	public void exportRepeatColumnTitles() {
		Book book = Util.loadBook(this, "book/blank.xlsx");
		
		File temp = Setup.getTempFile("Issue1035ExportBlank",".pdf");
		
		try {
			exportBook(book.getInternalBook(), temp);
			Assert.assertTrue("blank document should throw IOException", false);
		} catch(IOException ex) {
			Assert.assertTrue(ex.getMessage(), true);
		}
	}
	
	private void exportBook(SBook book, File file) throws IOException {
		
		PdfExporter exporter = new PdfExporter();
		exporter.export(book, file);
	}

}
