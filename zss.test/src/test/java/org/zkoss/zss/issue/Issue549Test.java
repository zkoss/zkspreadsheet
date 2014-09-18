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

public class Issue549Test {
	
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
	public void exportLargeTextA1WrapNoGridBottom() {
		Book book = Util.loadBook(this, "book/549-large-text-wrap-nogrid-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue549WrapNoGridBottomTest",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		Util.open(temp);
	}
	
	@Test
	public void exportLargeTextA1WrapGridBottom() {
		Book book = Util.loadBook(this, "book/549-large-text-wrap-grid-bottom-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue549WrapGridBottomTest",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		Util.open(temp);
	}

	@Test
	public void exportLargeTextA1WrapGridMiddle() {
		Book book = Util.loadBook(this, "book/549-large-text-wrap-grid-middle-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue549WrapGridMiddleTest",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		Util.open(temp);
	}
	
	@Test
	public void exportLargeTextA1WrapGridTop() {
		Book book = Util.loadBook(this, "book/549-large-text-wrap-grid-top-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue549WrapGridTopTest",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		Util.open(temp);
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
