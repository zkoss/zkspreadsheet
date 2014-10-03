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

public class Issue695Test {
	
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
	 * Test render sheet that fit to 1 page wide, 1 pages tall
	 */
	@Test
	public void exportFitPage1x1() { //1 page wide, 1 pages tall 
		Book book = Util.loadBook(this, "book/695-fit-page1x1-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue695FitPage1x1Test",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		Util.open(temp);
	}
	
	/**
	 * Test render sheet that fit to 1 page wide, 2 pages tall
	 */
	@Test
	public void exportFitPage1x2() { //1 page wide, 2 pages tall 
		Book book = Util.loadBook(this, "book/695-fit-page1x2-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue695FitPage1x2Test",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		Util.open(temp);
	}

	/**
	 * Test render sheet that fit to 2 page wide, 1 pages tall
	 */
	@Test
	public void exportFitPage2x1() { //2 page wide, 1 pages tall 
		Book book = Util.loadBook(this, "book/695-fit-page2x1-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue695FitPage2x1Test",".pdf");
		
		exportBook(book.getInternalBook(), temp);
		
		Util.open(temp);
	}
	
	/**
	 * Test render sheet that fit to 2 page wide, 2 pages tall
	 */
	@Test
	public void exportFitPage2x2() { //2 page wide, 2 pages tall 
		Book book = Util.loadBook(this, "book/695-fit-page2x2-pdf.xlsx");
		
		File temp = Setup.getTempFile("Issue695FitPage2x2Test",".pdf");
		
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
