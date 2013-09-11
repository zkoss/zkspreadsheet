package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.zkoss.zss.api.CellOperationUtil.applyFontBoldweight;
import static org.zkoss.zss.api.Ranges.range;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URISyntaxException;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.EditableCellStyle;
import org.zkoss.zss.api.model.EditableFont;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.Sheet;

/**
 * ZSS-408.
 * ZSS-414. ZSS-415.
 * ZSS-425. ZSS-426.
 * ZSS-435.
 * ZSS-439.
 * @author kuro
 *
 */
public class Issue400Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		ZssContext.setThreadLocal(new ZssContext(Locale.TAIWAN,-1));
	}
	
	@After
	public void tearDown() throws Exception {
		ZssContext.setThreadLocal(null);
	}

	@Test
	public void testZSS437() throws IOException, URISyntaxException {
		
		String EXPORT_FILE = "book/zss437.xlsx";
		
		Book workbook = Util.loadBook("book/blank.xlsx");
		Sheet sheet = workbook.getSheet("Sheet1");
		Range column0 = range(sheet, "A1");
		Util.export(workbook, EXPORT_FILE);
		
		column0.setCellEditText("Bold");
		applyFontBoldweight(column0, Font.Boldweight.BOLD);
		column0.setColumnWidth(100);
		assertEquals(100, sheet.getColumnWidth(0));
		
		//load saved file to validate width is saved
		Util.export(workbook, EXPORT_FILE);
		workbook = Util.loadBook(EXPORT_FILE);
		sheet = workbook.getSheet("Sheet1");
		assertEquals(100, sheet.getColumnWidth(0));
		//for human eye checking
		Util.open(EXPORT_FILE);
		
	}
	
	/**
	 * Endless loop when export demo swineFlu.xls pdf
	 */
	@Test
	public void testZSS426() throws IOException {
		
		final String filename = "book/426-swineFlu.xls";
		final InputStream is =  getClass().getResourceAsStream(filename);
		Book workbook = Importers.getImporter().imports(is, filename);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	exporter.export(workbook,fos);
    	
    	// import book again
    	Importers.getImporter().imports(temp,"test");

	}
	
	/**
	 * insert whole row or column when overlap merge cell with border style will cause unexpected result
	 */
	@Ignore
	@Test
	public void testZSS435() throws IOException {
		final String filename = "book/blank.xlsx";
		final InputStream is =  getClass().getResourceAsStream(filename);
		Book workbook = Importers.getImporter().imports(is, filename);
		
		Sheet sheet = workbook.getSheet("Sheet1");
		
		Range range = Ranges.range(sheet, "B2:D4");
		range.merge(false);
		
		Range columnC = Ranges.range(sheet, "C1");
		columnC.toColumnRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		assertEquals(BorderType.THIN, columnC.getCellStyle().getBorderBottom());
		
	}
	
	/**
	 * Cannot save 2003 format if the file contains auto filter configuration.
	 */
	@Test
	public void testZSS408() throws IOException {
		
		final String filename = "book/408-save-autofilter.xls";
		final InputStream is =  getClass().getResourceAsStream(filename);
		Book workbook = Importers.getImporter().imports(is, filename);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	exporter.export(workbook,fos);
    	
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
		Book workbook = Importers.getImporter().imports(is, filename);
		Sheet sheet1 = workbook.getSheet("Sheet1");
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
		Book workbook = Importers.getImporter().imports(is, filename);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	exporter.export(workbook, fos);
    	
    	// import book again
    	Importers.getImporter().imports(temp, "test");
	}
	
	@Test
	public void testZSS425() throws IOException {
		
		final String filename = "book/425-updateStyle.xlsx";
		final InputStream is = getClass().getResourceAsStream(filename);
		Book book = Importers.getImporter().imports(is, filename);
		
		Range r = Ranges.range(book.getSheetAt(0),0,0);
		
		CellOperationUtil.applyBackgroundColor(r, "#f0f000");
		r = Ranges.range(book.getSheetAt(0),0,0);
		Assert.assertEquals("#f0f000", r.getCellStyle().getBackgroundColor().getHtmlColor());
		
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	
    	java.io.File temp = java.io.File.createTempFile("test",".xlsx");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	//export first time
    	exporter.export(book, fos);
    	
		CellOperationUtil.applyBackgroundColor(r, "#00ff00");
		r = Ranges.range(book.getSheetAt(0),0,0);
		//change again
		Assert.assertEquals("#00ff00", r.getCellStyle().getBackgroundColor().getHtmlColor());
		
    	fos = new java.io.FileOutputStream(temp);
    	//export 2nd time
    	exporter.export(book, fos);
    	System.out.println(">>>write "+temp);
    	// import book again
    	book = Importers.getImporter().imports(temp, "test");
    	r = Ranges.range(book.getSheetAt(0),0,0);
    	//get #ff0000 if bug is not fixed
		Assert.assertEquals("#00ff00", r.getCellStyle().getBackgroundColor().getHtmlColor());
	}
	
	@Ignore("ZSS-432")
	@Test
	public void testZSS432() throws IOException {
		
		Book book = Util.loadBook("book/432.xlsx");
		
		Sheet sheet = Ranges.range(book.getSheetAt(0)).createSheet("newone");
		
		File temp = Setup.getTempFile();
		//export first time
		Exporters.getExporter().export(book, temp);
		
		sheet = book.getSheet("newone");
		
		//edit the new sheet
		Ranges.range(sheet,"A1").setCellEditText("ABC");
		
		//export again
		Exporters.getExporter().export(book, temp);
		
		//import, should't get any error
		Importers.getImporter().imports(temp, "test");
		
		sheet = book.getSheet("newone");
		
		//verify the value
		Assert.assertEquals("ABC", Ranges.range(sheet,"A1").getCellEditText());
	}
	
	
	@Test
	public void testZSS430() throws IOException {
		
		Book book = Util.loadBook("book/430-export-formula.xlsx");
		Sheet sheet = book.getSheet("formula-math");
		Assert.assertEquals("2.09", Ranges.range(sheet,"C7").getCellFormatText());
		
		
		File temp = Setup.getTempFile();
		//export first time
		Exporters.getExporter().export(book, temp);

		//import, should't get any error
		Importers.getImporter().imports(temp, "test");
		
		sheet = book.getSheet("formula-math");
		Assert.assertEquals("2.09", Ranges.range(sheet,"C7").getCellFormatText());
		
		System.out.println(">>>>"+temp);
		//can't be opened by excel
		//how to 
	}
	
	@Ignore("ZSS-429")
	@Test
	public void testZSS429() throws IOException {
		
		Book book = Util.loadBook("book/429-autofilter.xlsx");
		Sheet sheet = book.getSheetAt(0);
		Assert.assertEquals("A", Ranges.range(sheet,"C4").getCellFormatText());
		
		File temp = Setup.getTempFile();
		//export first time
		Exporters.getExporter().export(book, temp);
		
		//shouln't cause exception
		Ranges.range(sheet,"C4").toColumnRange().insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		
		Assert.assertEquals("A", Ranges.range(sheet,"D4").getCellFormatText());
		
		Exporters.getExporter().export(book, temp);
		
		
		//import, should't get any error
		Importers.getImporter().imports(temp, "test");
				
		sheet = book.getSheetAt(0);
		Assert.assertEquals("A", Ranges.range(sheet,"D4").getCellFormatText());
		
	}
	
	@Ignore("ZSS-431")
	@Test
	public void testZSS431() throws IOException {
		testZSS431_0(Util.loadBook("book/blank.xlsx"));
		testZSS431_0(Util.loadBook("book/blank.xls"));
		testZSS431_0(Util.loadBook("book/431-freezepanel.xlsx"));
	}
	
	public void testZSS431_0(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		Assert.assertEquals(0, sheet.getRowFreeze());
		Assert.assertEquals(0, sheet.getColumnFreeze());
		
		Ranges.range(sheet).setFreezePanel(7, 5);
		Assert.assertEquals(7, sheet.getRowFreeze());
		Assert.assertEquals(5, sheet.getColumnFreeze());
		
		File temp = Setup.getTempFile();
		//export first time
		Exporters.getExporter().export(book, temp);
		

		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(7, sheet.getRowFreeze());
		Assert.assertEquals(5, sheet.getColumnFreeze());
		
		//TODO how to verify it in Excel?
		
		
	}

	@Test
	public void testZSS439() throws IOException {
		
		final String filename = "book/439-rows.xlsx";
		final InputStream is = getClass().getResourceAsStream(filename);
		Book book = Importers.getImporter().imports(is, filename);
		
		// fill text on 3 rows
		Sheet sheet = book.getSheetAt(0);
		Ranges.range(sheet, "A1:A3").setCellEditText("test");
		
		// apply bold style on such rows
		Range r = Ranges.range(sheet, "A1:A3");
		EditableFont f = r.getCellStyleHelper().createFont(null);
		f.setBoldweight(org.zkoss.zss.api.model.Font.Boldweight.BOLD);
		EditableCellStyle s = r.getCellStyleHelper().createCellStyle(null);
		s.setFont(f);
		r.setCellStyle(s);
		
		// remove two rows
		Ranges.range(sheet, "1:2").toRowRange().delete(DeleteShift.UP);
		
		// read the style, it should not occur exception.
		Ranges.range(sheet, "A1").getCellStyle();
	}
}
