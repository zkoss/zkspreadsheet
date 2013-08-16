package org.zkoss.zss.api.impl;

import java.io.IOException;
import java.io.InputStream;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

/**
 * ZSS-408.
 * ZSS-414.
 * ZSS-415.
 * @author kuro
 *
 */
public class Issue400Test {
	
	private static Book _workbook;
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	/**
	 * Cannot save 2003 format if the file contains auto filter configuration.
	 */
	@Test
	public void testZSS408() throws IOException {
		
		final String filename = "book/408-save-autofilter.xls";
		final InputStream is =  getClass().getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	exporter.export(_workbook,fos);
    	
    	// import book again
    	Importers.getImporter().imports(temp,"test");

	}
	
	/**
	 * 1. load a blank sheet (excel 2003)
	 * 2. set cell empty string to any cell will cause exception
	 */
	@Test
	public void testZSS414() throws IOException {
		final String filename = "book/blank.xls";
		final InputStream is =  getClass().getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Range r = Ranges.range(sheet1, "C1");
		r.setCellEditText("");
	}
	
	/**
	 * 1. import a book with comment
	 * 2. export
	 * 3. import again will throw exception
	 */
	@Test
	public void testZSS415() throws IOException {
		
		final String filename = "book/415-commentUnsupport.xls";
		final InputStream is =  getClass().getResourceAsStream(filename);
		_workbook = Importers.getImporter().imports(is, filename);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	exporter.export(_workbook, fos);
    	
    	// import book again
    	Importers.getImporter().imports(temp, "test");
	}
	
	
//	@Test //not fix yet
	public void testZSS425() throws IOException {
		
		final String filename = "book/425-updateStyle.xlsx";
		final InputStream is = getClass().getResourceAsStream(filename);
		Book book = Importers.getImporter().imports(is, filename);
		
		Range r = Ranges.range(book.getSheetAt(0),0,0);
		
		CellOperationUtil.applyBackgroundColor(r, "#FF0000");
		r = Ranges.range(book.getSheetAt(0),0,0);
		Assert.assertEquals("#ff0000", r.getCellStyle().getBackgroundColor().getHtmlColor());
		
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	//export first time
    	exporter.export(book, fos);
    	
		CellOperationUtil.applyBackgroundColor(r, "#0000FF");
		r = Ranges.range(book.getSheetAt(0),0,0);
		//change again
		Assert.assertEquals("#0000ff", r.getCellStyle().getBackgroundColor().getHtmlColor());
		
    	fos = new java.io.FileOutputStream(temp);
    	//export 2nd time
    	exporter.export(book, fos);
    	
    	// import book again
    	book = Importers.getImporter().imports(temp, "test");
    	r = Ranges.range(book.getSheetAt(0),0,0);
    	//get #ff0000 if bug is not fixed
		Assert.assertEquals("#0000ff", r.getCellStyle().getBackgroundColor().getHtmlColor());
	}


}
