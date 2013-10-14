package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;
import static org.zkoss.zss.api.CellOperationUtil.applyFontBoldweight;
import static org.zkoss.zss.api.Ranges.range;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URISyntaxException;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.poi.xssf.usermodel.XSSFComment;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.zss.POIUtil;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.BookSeriesBuilder;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Exporter;
import org.zkoss.zss.api.Exporters;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Range.DeleteShift;
import org.zkoss.zss.api.Range.InsertCopyOrigin;
import org.zkoss.zss.api.Range.InsertShift;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.CellData.CellType;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.EditableCellStyle;
import org.zkoss.zss.api.model.EditableFont;
import org.zkoss.zss.api.model.Font;
import org.zkoss.zss.api.model.Hyperlink;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zssex.ui.impl.XUtils;

/**
 * ZSS-408.
 * ZSS-414. ZSS-415.
 * ZSS-418.
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
		Setup.pushZssContextLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssContextLocale();
	}
	
	@Test
	public void testZSS400_2007() throws IOException {
		testZSS400HyperlinkShift0(Util.loadBook(this, "book/blank.xlsx"));
		testZSS400HyperlinkShiftColumn(Util.loadBook(this, "book/400-shift.xlsx"));
		testZSS400HyperlinkShiftRow(Util.loadBook(this, "book/400-shift.xlsx"));
		testZSS400HyperlinkShiftBoth(Util.loadBook(this, "book/400-shift.xlsx"));
		
	}
	
	public void testZSS400HyperlinkShift0(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		Range r1 = Ranges.range(sheet, "A1");
		Range r2 = Ranges.range(sheet, "B1");
		Range r3 = Ranges.range(sheet, "C1");
		
		r1.setCellEditText("AAA");
		r2.setCellHyperlink(HyperlinkType.URL, "http://www.zkoss.org", "www.zkoss.org");
		
		Assert.assertEquals("AAA", r1.getCellFormatText());
		Assert.assertEquals("www.zkoss.org", r2.getCellFormatText());
		Assert.assertEquals("http://www.zkoss.org", r2.getCellHyperlink().getAddress());
		Assert.assertEquals("www.zkoss.org", r2.getCellHyperlink().getLabel());
		Assert.assertEquals(HyperlinkType.URL, r2.getCellHyperlink().getType());
		
		r2.toColumnRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		
		Assert.assertEquals("AAA", r1.getCellFormatText());
		Assert.assertEquals("", r2.getCellFormatText());
		Assert.assertEquals(null, r2.getCellHyperlink());
		
		r2.setCellEditText("BBB");
		Assert.assertEquals("BBB", r2.getCellFormatText());
		Assert.assertNull(r2.getCellHyperlink());
		
		Assert.assertEquals("www.zkoss.org", r3.getCellFormatText());
		Assert.assertEquals("www.zkoss.org", r3.getCellHyperlink().getLabel());
		Assert.assertEquals("http://www.zkoss.org", r3.getCellHyperlink().getAddress());
		Assert.assertEquals(HyperlinkType.URL, r3.getCellHyperlink().getType());
		
		
		r3.setCellEditText("CCC");
		Assert.assertEquals("CCC", r3.getCellFormatText());
		Assert.assertEquals("CCC", r3.getCellHyperlink().getLabel());
		Assert.assertEquals("http://www.zkoss.org", r3.getCellHyperlink().getAddress());
		Assert.assertEquals(HyperlinkType.URL, r3.getCellHyperlink().getType());
		
		
		r2.toColumnRange().delete(DeleteShift.DEFAULT);
		Assert.assertEquals("CCC", r2.getCellFormatText());
		Assert.assertEquals("CCC", r2.getCellHyperlink().getLabel());
		Assert.assertEquals("http://www.zkoss.org", r2.getCellHyperlink().getAddress());
		Assert.assertEquals(HyperlinkType.URL, r2.getCellHyperlink().getType());
		
		
		r2.toColumnRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		Assert.assertEquals("AAA", r1.getCellFormatText());
		Assert.assertEquals("", r2.getCellFormatText());
		Assert.assertEquals(null, r2.getCellHyperlink());
		
		r2.setCellEditText("DDD");
		Assert.assertEquals("DDD", r2.getCellFormatText());
		Assert.assertNull(r2.getCellHyperlink());
		
		Assert.assertEquals("CCC", r3.getCellFormatText());
		Assert.assertEquals("CCC", r3.getCellHyperlink().getLabel());
		Assert.assertEquals("http://www.zkoss.org", r3.getCellHyperlink().getAddress());
		Assert.assertEquals(HyperlinkType.URL, r3.getCellHyperlink().getType());
		
		book = Util.swap(book);
		
		sheet = book.getSheetAt(0);
		r1 = Ranges.range(sheet, "A1");
		r2 = Ranges.range(sheet, "B1");
		r3 = Ranges.range(sheet, "C1");
		
		Assert.assertEquals("AAA", r1.getCellFormatText());
		Assert.assertEquals("DDD", r2.getCellFormatText());
		Assert.assertNull(r2.getCellHyperlink());
		Assert.assertEquals("CCC", r3.getCellFormatText());
		Assert.assertEquals("CCC", r3.getCellHyperlink().getLabel());
		Assert.assertEquals("http://www.zkoss.org", r3.getCellHyperlink().getAddress());
		Assert.assertEquals(HyperlinkType.URL, r3.getCellHyperlink().getType());
		
	}
	
	public void testZSS400HyperlinkShiftColumn(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
		
		Ranges.range(sheet,"D1:E1").toColumnRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"H2").getCellHyperlink());
		Assert.assertNull(Ranges.range(sheet,"H7").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"J2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"J7").getCellHyperlink().getAddress());
		
		Ranges.range(sheet,"D1:E3").insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_NONE);
		
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"J2").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"J7").getCellHyperlink().getAddress());
		
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"L2").getCellHyperlink().getAddress());
		
		
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"J2").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"J7").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"L2").getCellHyperlink().getAddress());
		
		
		
		Ranges.range(sheet,"E1:F3").delete(DeleteShift.LEFT);
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"L2").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"J2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"J7").getCellHyperlink().getAddress());
		
		
		Ranges.range(sheet,"E1:F1").toColumnRange().delete(DeleteShift.LEFT);
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"J2").getCellHyperlink());
		Assert.assertNull(Ranges.range(sheet,"J7").getCellHyperlink());
		
		
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"J2").getCellHyperlink());
		Assert.assertNull(Ranges.range(sheet,"J7").getCellHyperlink());
	}
	
	public void testZSS400HyperlinkShiftRow(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
		
		Ranges.range(sheet,"A3:A4").toRowRange().insert(InsertShift.DEFAULT, InsertCopyOrigin.FORMAT_NONE);
		
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"A7").getCellHyperlink());
		Assert.assertNull(Ranges.range(sheet,"H7").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B9").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H9").getCellHyperlink().getAddress());
		
		Ranges.range(sheet,"A3:C4").insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_NONE);
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"B9").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B11").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H9").getCellHyperlink().getAddress());
		
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"B9").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B11").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H9").getCellHyperlink().getAddress());
		
		
		Ranges.range(sheet,"A3:C4").delete(DeleteShift.UP);
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"B11").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B9").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H9").getCellHyperlink().getAddress());
		
		Ranges.range(sheet,"A3:C4").toRowRange().delete(DeleteShift.UP);
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"B9").getCellHyperlink());
		Assert.assertNull(Ranges.range(sheet,"H9").getCellHyperlink());
		
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
		Assert.assertNull(Ranges.range(sheet,"B9").getCellHyperlink());
		Assert.assertNull(Ranges.range(sheet,"H9").getCellHyperlink());
		
	}
	
	public void testZSS400HyperlinkShiftBoth(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"B2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/B7", Ranges.range(sheet,"B7").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
		
		Ranges.range(sheet,"A1:C3").shift(4, 1);
		Assert.assertNull(Ranges.range(sheet,"B2").getCellHyperlink());
		Assert.assertNull(Ranges.range(sheet,"B7").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"C6").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
		
		book = Util.swap(book);
		sheet = book.getSheetAt(0);
		Assert.assertNull(Ranges.range(sheet,"B2").getCellHyperlink());
		Assert.assertNull(Ranges.range(sheet,"B7").getCellHyperlink());
		Assert.assertEquals("http://sheet.test/B2", Ranges.range(sheet,"C6").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H2", Ranges.range(sheet,"H2").getCellHyperlink().getAddress());
		Assert.assertEquals("http://sheet.test/H7", Ranges.range(sheet,"H7").getCellHyperlink().getAddress());
	}
	
	@Test
	public void testZSS447_2003() throws IOException {
		testZSS447(Util.loadBook(this, "book/447-dateFormatDisplay.xls"));
	}
	
	@Test
	public void testZSS447_2007() throws IOException {
		testZSS447(Util.loadBook(this, "book/447-dateFormatDisplay.xlsx"));
	}
	
	public void testZSS447(Book book) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		Range a1 = Ranges.range(sheet, "A1");
		Range a2 = Ranges.range(sheet, "A2");
		Range a3 = Ranges.range(sheet, "A3");
		Range a4 = Ranges.range(sheet, "A4");
		
		Setup.pushZssContextLocale(Locale.TAIWAN);
		try{
			org.junit.Assert.assertEquals("yyyy/m/d", a1.getCellStyle().getDataFormat()); //default format will depends on LOCAL
			org.junit.Assert.assertEquals("2013/12/24", a1.getCellFormatText());
			
			org.junit.Assert.assertEquals("yyyy/m/d", a2.getCellStyle().getDataFormat());
			org.junit.Assert.assertEquals("2013/12/24", a2.getCellFormatText());
			
			org.junit.Assert.assertEquals("m/d/yyyy", a3.getCellStyle().getDataFormat()); // PASS is POI
			org.junit.Assert.assertEquals("12/24/2013", a3.getCellFormatText());// PASS is POI
		}finally{
			Setup.popZssContextLocale();
		}
		Setup.pushZssContextLocale(Locale.US);
		try{
			org.junit.Assert.assertEquals("m/d/yyyy", a1.getCellStyle().getDataFormat()); //this cell contains default format, it should depends on locale
			org.junit.Assert.assertEquals("12/24/2013", a1.getCellFormatText());
			
			org.junit.Assert.assertEquals("m/d/yyyy", a2.getCellStyle().getDataFormat()); //this cell contains default format, it should depends on locale
			org.junit.Assert.assertEquals("12/24/2013", a2.getCellFormatText());
			
			org.junit.Assert.assertEquals("m/d/yyyy", a3.getCellStyle().getDataFormat()); //this is custom format, it regardless locale.
			org.junit.Assert.assertEquals("12/24/2013", a3.getCellFormatText());
		}finally{
			Setup.popZssContextLocale();
		}
		
		//test input, depends on locale
		Range a5 = Ranges.range(sheet, "A5");
		Range a6 = Ranges.range(sheet, "A6");
		Setup.pushZssContextLocale(Locale.TAIWAN);
		try{
			org.junit.Assert.assertEquals("General", a5.getCellStyle().getDataFormat());
			org.junit.Assert.assertEquals("General", a6.getCellStyle().getDataFormat());
			
			a5.setCellEditText("2013/2/1");
			org.junit.Assert.assertEquals(CellType.NUMERIC, a5.getCellData().getType());
			org.junit.Assert.assertEquals("yyyy/m/d", a5.getCellStyle().getDataFormat()); //default format will depends on LOCAL
			org.junit.Assert.assertEquals("2013/2/1", a5.getCellFormatText());
			
			a6.setCellEditText("2/1/2013");
			org.junit.Assert.assertEquals(CellType.STRING, a6.getCellData().getType());
			org.junit.Assert.assertEquals("General", a6.getCellStyle().getDataFormat());
			org.junit.Assert.assertEquals("2/1/2013", a6.getCellFormatText());
		}finally{
			Setup.popZssContextLocale();
		}
		
		Range a7 = Ranges.range(sheet, "A7");
		Range a8 = Ranges.range(sheet, "A8");
		Setup.pushZssContextLocale(Locale.US);
		try{
			org.junit.Assert.assertEquals("General", a7.getCellStyle().getDataFormat());
			org.junit.Assert.assertEquals("General", a8.getCellStyle().getDataFormat());
			
			a7.setCellEditText("2/1/2013");
			org.junit.Assert.assertEquals(CellType.NUMERIC, a7.getCellData().getType());
			org.junit.Assert.assertEquals("m/d/yyyy", a7.getCellStyle().getDataFormat());
			org.junit.Assert.assertEquals("2/1/2013", a7.getCellFormatText());
			
			a8.setCellEditText("2013/2/1");
			org.junit.Assert.assertEquals(CellType.STRING, a8.getCellData().getType());
			org.junit.Assert.assertEquals("General", a8.getCellStyle().getDataFormat()); //default format will depends on LOCAL
			org.junit.Assert.assertEquals("2013/2/1", a8.getCellFormatText());
		}finally{
			Setup.popZssContextLocale();
		}

	}

	@Test
	public void testZSS437() throws IOException, URISyntaxException {
		
		File target = Setup.getTempFile("zss437",".xlsx");
		
		Book workbook = Util.loadBook(this,"book/blank.xlsx");
		Sheet sheet = workbook.getSheet("Sheet1");
		Range column0 = range(sheet, "A1");
		Util.export(workbook, target);
		
		column0.setCellEditText("Bold");
		applyFontBoldweight(column0, Font.Boldweight.BOLD);
		column0.setColumnWidth(100);
		assertEquals(100, sheet.getColumnWidth(0));
		
		//load saved file to validate width is saved
		Util.export(workbook, target);
		workbook = Util.loadBook(target);
		sheet = workbook.getSheet("Sheet1");
		assertEquals(100, sheet.getColumnWidth(0));
		//for human eye checking
		Util.open(target);
		
	}
	
	/**
	 * Endless loop when export demo swineFlu.xls pdf
	 */
	@Test
	public void testZSS426() throws IOException {
		
		final String filename = "book/426-swineFlu.xls";
		Book book = Util.loadBook(this, filename);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	exporter.export(book,fos);
    	
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
		Book book = Util.loadBook(this, filename);
		
		Sheet sheet = book.getSheet("Sheet1");
		
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
		Book book = Util.loadBook(this, filename);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	exporter.export(book,fos);
    	
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
		Book book = Util.loadBook(this, filename);
		Sheet sheet1 = book.getSheet("Sheet1");
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
		Book book = Util.loadBook(this, filename);
		
		// export work book
    	Exporter exporter = Exporters.getExporter();
    	java.io.File temp = java.io.File.createTempFile("test",".xls");
    	java.io.FileOutputStream fos = new java.io.FileOutputStream(temp);
    	exporter.export(book, fos);
    	
    	// import book again
    	Importers.getImporter().imports(temp, "test");
	}
	
	@Test
	public void testZSS425() throws IOException {
		
		final String filename = "book/425-updateStyle.xlsx";
		Book book = Util.loadBook(this, filename);
		
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
	
	@Test
	public void testZSS432() throws IOException {
		
//		Book book = Util.loadBook(this,"book/blank.xlsx");
		Book book = Util.loadBook(this,"book/432.xlsx");
		
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
		
		Book book = Util.loadBook(this,"book/430-export-formula.xlsx");
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
	
	@Test
	public void testZSS429() throws IOException {
		
		Book book = Util.loadBook(this,"book/429-autofilter.xlsx");
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

	@Test 
	public void testZSS431() throws IOException {
		testZSS431_0(Util.loadBook(this,"book/431-freezepanel.xlsx"),7,4);
		testZSS431_0(Util.loadBook(this,"book/431-freezepanel.xlsx"),7,0);
		testZSS431_0(Util.loadBook(this,"book/431-freezepanel.xlsx"),0,4);
		testZSS431_0(Util.loadBook(this,"book/431-freezepanel.xls"),7,4);
		testZSS431_0(Util.loadBook(this,"book/431-freezepanel.xls"),7,0);
		testZSS431_0(Util.loadBook(this,"book/431-freezepanel.xls"),0,4);
	}
	
	public void testZSS431_0(Book book,int row, int column) throws IOException {
		Sheet sheet = book.getSheetAt(0);
		Assert.assertEquals(0, sheet.getRowFreeze());
		Assert.assertEquals(0, sheet.getColumnFreeze());
		
		Ranges.range(sheet).setFreezePanel(row, column);
		Assert.assertEquals(row, sheet.getRowFreeze());
		Assert.assertEquals(column, sheet.getColumnFreeze());
		
		Assert.assertEquals(0, sheet.getPoiSheet().getTopRow());
		Assert.assertEquals(0, sheet.getPoiSheet().getLeftCol());
		
		File temp = Setup.getTempFile();
		//export first time
		Exporters.getExporter().export(book, temp);
		

		book = Importers.getImporter().imports(temp, "test");
		sheet = book.getSheetAt(0);
		
		Assert.assertEquals(row, sheet.getRowFreeze());
		Assert.assertEquals(column, sheet.getColumnFreeze());
		Assert.assertEquals(0, sheet.getPoiSheet().getTopRow());
		Assert.assertEquals(0, sheet.getPoiSheet().getLeftCol());
		
		//TODO how to verify it in Excel?
	}

	@Test
	public void testZSS439() throws IOException {
		
		final String filename = "book/439-rows.xlsx";
		Book book = Util.loadBook(this, filename);
		
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
	
	@Test
	public void testZSS412() throws IOException {
		
		Book book = Util.loadBook(this,"book/412-overlap-undo.xls");
		Sheet sheet = book.getSheetAt(0);
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("A2:A3", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Range r1 = Ranges.range(sheet,"A1:A3");
		Range r2 = Ranges.range(sheet,"A2");
		r1.paste(r2);	
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("A3:A4", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Range r3 = Ranges.range(sheet,"A2:A3");
		r3.merge(false);
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("A2:A3", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		r3 = Ranges.range(sheet,"A2:B3");
		r3.merge(true);
		Assert.assertEquals(2, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("A2:B2", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		Assert.assertEquals("A3:B3", sheet.getPoiSheet().getMergedRegion(1).formatAsString());
	}
	@Test
	public void testZSS412_395_1() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		Sheet sheet = book.getSheetAt(0);
		Range rangeA = Ranges.range(sheet, "H11:J13");
		rangeA.merge(false); // merge a 3 x 3
		
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("H11:J13", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		Range rangeB = Ranges.range(sheet, "H12"); // a whole row cross the merged cell
		rangeB.toRowRange().unmerge(); // perform unmerge operation
	
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
		
		//again on column
		rangeA = Ranges.range(sheet, "H11:J13");
		rangeA.merge(false); // merge a 3 x 3
		Assert.assertEquals(1, sheet.getPoiSheet().getNumMergedRegions());
		Assert.assertEquals("H11:J13", sheet.getPoiSheet().getMergedRegion(0).formatAsString());
		
		rangeB = Ranges.range(sheet, "H12"); // a whole row cross the merged cell
		rangeB.toColumnRange().unmerge(); // perform unmerge operation
		Assert.assertEquals(0, sheet.getPoiSheet().getNumMergedRegions());
	}
	
	@Test
	public void testZSS418() throws IOException {
		Book book = Util.loadBook(this,"book/418-comment.xlsx");
		Sheet sheet = book.getSheetAt(0);
		XSSFSheet ps = (XSSFSheet)sheet.getPoiSheet();
		
		// check original comments
		String[] refs = {"C3", "D3", "E3", "F3", "H3", "I3", "C4", "C5", "C6", "C8", "C9"};
		String[] txts = {"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10"}; // null indicates test comment not existed
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "G3").delete(DeleteShift.LEFT);
		refs = new String[]{"C3", "D3", "E3", "F3", "G3", "H3", "C4", "C5", "C6", "C8", "C9", "I3"};
		txts = new String[]{"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null};
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "C7").delete(DeleteShift.UP);
		refs = new String[]{"C3", "D3", "E3", "F3", "G3", "H3", "C4", "C5", "C6", "C7", "C8", "C9"};
		txts = new String[]{"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null};
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "F3").insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		refs = new String[]{"C3", "D3", "E3", "G3", "H3", "I3", "C4", "C5", "C6", "C7", "C8", "F3"};
		txts = new String[]{"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null};
		testZSS418_0(ps, refs, txts);
		
		Ranges.range(sheet, "C6").insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		refs = new String[]{"C3", "D3", "E3", "G3", "H3", "I3", "C4", "C5", "C7", "C8", "C9", "C6"};
		txts = new String[]{"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null};
		testZSS418_0(ps, refs, txts);
		
		Ranges.range(sheet, "F").toColumnRange().delete(DeleteShift.LEFT);
		refs = new String[]{"C3", "D3", "E3", "F3", "G3", "H3", "C4", "C5", "C7", "C8", "C9", "I3"};
		txts = new String[]{"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null};
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "6").toRowRange().delete(DeleteShift.UP);
		refs = new String[]{"C3", "D3", "E3", "F3", "G3", "H3", "C4", "C5", "C6", "C7", "C8", "C9"};
		txts = new String[]{"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null};
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "G").toColumnRange().insert(InsertShift.RIGHT, InsertCopyOrigin.FORMAT_RIGHT_BELOW);
		refs = new String[]{"C3", "D3", "E3", "F3", "H3", "I3", "C4", "C5", "C6", "C7", "C8", "G3"};
		txts = new String[]{"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null};
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "7").toRowRange().insert(InsertShift.DOWN, InsertCopyOrigin.FORMAT_LEFT_ABOVE);
		refs = new String[]{"C3", "D3", "E3", "F3", "H3", "I3", "C4", "C5", "C6", "C8", "C9", "C7"};
		txts = new String[]{"c0", "c1", "c2", "c3", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null};
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "H3:I3").shift(0, -2);
		refs = new String[]{"C3", "D3", "E3", "F3", "G3", "C4", "C5", "C6", "C8", "C9", "H3", "I3"};
		txts = new String[]{"c0", "c1", "c2", "c4", "c5", "c6", "c7", "c8", "c9", "c10", null, null};
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "C8:C9").shift(-2, 0);
		refs = new String[]{"C3", "D3", "E3", "F3", "G3", "C4", "C5", "C6", "C7", "C8", "C9"};
		txts = new String[]{"c0", "c1", "c2", "c4", "c5", "c6", "c7", "c9", "c10", null, null};
		testZSS418_0(ps, refs, txts);

		Ranges.range(sheet, "C3:G7").shift(-1, -1);
		refs = new String[]{"C3", "D3", "E3", "F3", "G3", "C4", "C5", "C6", "C7"};
		txts = new String[]{null, null, null, null, null, null, null, null, null};
		testZSS418_0(ps, refs, txts);
		refs = new String[]{"B2", "C2", "D2", "E2", "F2", "B3", "B4", "B5", "B6"};
		txts = new String[]{"c0", "c1", "c2", "c4", "c5", "c6", "c7", "c9", "c10"};
		testZSS418_0(ps, refs, txts);
		
		Ranges.range(sheet, "B2:F6").shift(2, 2);
		refs = new String[]{"B2", "C2", "D2", "E2", "F2", "B3", "B4", "B5", "B6"};
		txts = new String[]{null, null, null, null, null, null, null, null, null};
		testZSS418_0(ps, refs, txts);
		refs = new String[]{"D4", "E4", "F4", "G4", "H4", "D5", "D6", "D7", "D8"};
		txts = new String[]{"c0", "c1", "c2", "c4", "c5", "c6", "c7", "c9", "c10"};
		testZSS418_0(ps, refs, txts);
	}
	
	public void testZSS418_0(XSSFSheet sheet, String[] refs, String[] texts) {
		for(int i = 0; i < refs.length; ++i) {
			XSSFComment comment = POIUtil.getComment(sheet, refs[i]);
			String text = texts[i];
			if(text != null) {
				Assert.assertNotNull(comment);
				Assert.assertEquals(texts[i], comment.getString().getString());
			} else {
				Assert.assertNull(comment);
			}
		}
	}
	
	@Test
	public void testZSS446() throws Exception {
		// load book
		Book book = Util.loadBook(this,"book/446-border.xlsx");
		

		// print setting >> with grid lines 
		Sheet sheet = book.getSheetAt(0);
		sheet.getPoiSheet().setPrintGridlines(true);
		
		File temp = Setup.getTempFile("zss446","pdf");
		
		Exporter pdfExporter = Exporters.getExporter("pdf");
		pdfExporter.export(book, temp);

		System.out.println(">>export pdf to "+temp);
		Util.open(temp);
	}

	/**
	 * reference to a JavaBean is not evaluated to a cell reference even the value is "=B2".
	 * So it's not related to this issue.
	 */
	@Test
	public void testZSS441(){
		Book workbook = Util.loadBook(this,"book/441-sum.xls");
		Sheet sheet = workbook.getSheet("issue");
		//sum an array row reference
		Range sumResult = Ranges.range(sheet, "B3");
		assertEquals(11, sumResult.getCellData().getDoubleValue().intValue());
		sumResult = Ranges.range(sheet, "C3");
		assertEquals(22, sumResult.getCellData().getDoubleValue().intValue());
		sumResult = Ranges.range(sheet, "D3");
		assertEquals(33, sumResult.getCellData().getDoubleValue().intValue());
	}
	
	/**
	 * "array formula" is the feature related to ZSS-441
	 * http://office.microsoft.com/en-us/excel-help/guidelines-and-examples-of-array-formulas-HA010228458.aspx
	 */
	@Test
	public void testZSS441ArrayFormula(){
		Book workbook = Util.loadBook(this,"book/441-sum.xls");
		Sheet sheet = workbook.getSheet("array");
		//array formula
		Range multiplyResult = Ranges.range(sheet, "C3");
		assertEquals(4, multiplyResult.getCellData().getDoubleValue().intValue());

		multiplyResult = Ranges.range(sheet, "C4");
		assertEquals(10, multiplyResult.getCellData().getDoubleValue().intValue());
		
		multiplyResult = Ranges.range(sheet, "C5");
		assertEquals(18, multiplyResult.getCellData().getDoubleValue().intValue());
		
		//array formula - column
		Range arrayColumn = Ranges.range(sheet, "B8");
		assertEquals(8, arrayColumn.getCellData().getDoubleValue().intValue());
		arrayColumn = Ranges.range(sheet, "B9");
		assertEquals(9, arrayColumn.getCellData().getDoubleValue().intValue());
		arrayColumn = Ranges.range(sheet, "B10");
		assertEquals(10, arrayColumn.getCellData().getDoubleValue().intValue());
		
		sheet = workbook.getSheet("issue");
		//array formula - row
		Range arrayRow = Ranges.range(sheet, "B2");
		assertEquals(1, arrayRow.getCellData().getDoubleValue().intValue());
		arrayRow = Ranges.range(sheet, "C2");
		assertEquals(2, arrayRow.getCellData().getDoubleValue().intValue());
		arrayRow = Ranges.range(sheet, "D2");
		assertEquals(3, arrayRow.getCellData().getDoubleValue().intValue());
	}
	
	@Test
	public void testZSS441ThreeDimensionalReference(){
		Book workbook = Util.loadBook(this,"book/441-sum.xls");
		Sheet sheet = workbook.getSheet("3D");
		Range sumResult = Ranges.range(sheet, "B1");
		assertEquals(11, sumResult.getCellData().getDoubleValue().intValue());
		sumResult = Ranges.range(sheet, "C1");
		assertEquals(22, sumResult.getCellData().getDoubleValue().intValue());
		sumResult = Ranges.range(sheet, "D1");
		assertEquals(33, sumResult.getCellData().getDoubleValue().intValue());
	}
	
	@Test
	public void testZSS441ExternalBook(){
		Book workbook = Util.loadBook(this,"book/441-sum.xls");
		Book anotherWorkbook = Util.loadBook(this,"book/441-another.xls");
		BookSeriesBuilder.getInstance().buildBookSeries(new Book[]{workbook, anotherWorkbook});
		
		Sheet sheet = workbook.getSheet("external");
		Range sumResult = Ranges.range(sheet, "A1");
		assertEquals(60, sumResult.getCellData().getDoubleValue().intValue());
	}
	
	@Test
	public void testZSS441FormulaInValidation(){
		Book workbook = Util.loadBook(this,"book/441-validation.xlsx");
		Sheet sheet = workbook.getSheet("validation");
		//A ~ D
		Range validationCell = Ranges.range(sheet, "B3");
		assertEquals(true,validationCell.getCellData().validateEditText("B"));
		assertEquals(false,validationCell.getCellData().validateEditText("E"));
		
		// only A
		validationCell = Ranges.range(sheet, "B4");
		assertEquals(true, validationCell.getCellData().validateEditText("A"));
		assertEquals(false, validationCell.getCellData().validateEditText("E"));
		
		// 2nd reference, only B
		validationCell = Ranges.range(sheet, "B5");
		assertEquals(true, validationCell.getCellData().validateEditText("B"));
		assertEquals(false,validationCell.getCellData().validateEditText("E"));
	}

	@Test
	public void testZSS401(){
		Book workbook = Util.loadBook(Issue400Test.class, "book/401-cut-merged.xlsx");
		Range srcRange = Ranges.range(workbook.getSheet("source"), "A1:C1");
		Range destRange = Ranges.range(workbook.getSheet("destination"),"A2:C2");
		
		CellOperationUtil.cut(srcRange, destRange);
		assertEquals(false, srcRange.hasMergedCell());
	}	
	
	@Test
	public void testZSS401CutInSameSheet(){
		Book workbook = Util.loadBook(Issue400Test.class, "book/401-cut-merged.xlsx");
		Range srcRange = Ranges.range(workbook.getSheet("source"), "A1:C1");
		Range destRange = Ranges.range(workbook.getSheet("source"),"A2:C2");
		
		CellOperationUtil.cut(srcRange, destRange);
		assertEquals(false, srcRange.hasMergedCell());
	}
	
	@Test
	public void testZSS453(){
		Book workbook = Util.loadBook(Issue400Test.class, "book/453-emptyHyperlink.xlsx");
		String correctHyperlink = "<a zs.t=\"SHyperlink\" z.t=\"1\" href=\"javascript:\" z.href=\"http://www.zkoss.org/\">zkoss</a>";
		assertEquals(correctHyperlink, XUtils.getRichCellHtmlText((XSheet)workbook.getPoiBook().getSheetAt(0),0,0));
		try{
			XUtils.getRichCellHtmlText((XSheet)workbook.getPoiBook().getSheetAt(0),1,1);
		}catch(NullPointerException npe){
			fail("empty hyperlink shouldn't throw an exception.");
		}
	}
	
	@Test
	public void testZSS456(){
		Book workbook = Util.loadBook(Issue400Test.class, "book/blank.xlsx");
		Sheet sheet = workbook.getSheetAt(0);
		Range target1 = Ranges.range(sheet, "B2:C3");
		Range target2 = Ranges.range(sheet, "A1:D4");
		target1.merge(false);
		Assert.assertTrue(target1.isMergedCell());
		Assert.assertFalse(target2.isMergedCell());
		Assert.assertTrue(target2.hasMergedCell());
		target1.unmerge();
		Assert.assertFalse(target1.isMergedCell());
		Assert.assertFalse(target2.hasMergedCell());
		Assert.assertFalse(target2.hasMergedCell());
	}
	
	@Test
	public void testZSS427(){
		Book workbook = Util.loadBook(Issue400Test.class, "book/427-export.xlsx");
		File file = Setup.getTempFile("zss427",".pdf");
		try{
			FileOutputStream fos = new FileOutputStream(file);
			Exporter pdfExporter = Exporters.getExporter("pdf");
			pdfExporter.export(workbook, fos);
		}catch(NullPointerException e){
			e.printStackTrace();
			fail("export failed for ZSS-427");
		}catch(Exception e){
			e.printStackTrace();
		}
	}
	
	@Test
	public void testZSS461(){
		Book book = Util.loadBook(Issue400Test.class, "book/blank.xlsx");
		Sheet sheet = book.getSheetAt(0);
		Ranges.range(sheet, "A1").setCellEditText("http://www.google.com");
		Ranges.range(sheet, "A2").setCellEditText("http://www.zkoss.org");
		
		Hyperlink link = Ranges.range(sheet, "A1").getCellHyperlink();
		Assert.assertNotNull(link);
		Assert.assertEquals(Hyperlink.HyperlinkType.URL,link.getType());
		Assert.assertEquals(link.getAddress(), "http://www.google.com");
		Assert.assertEquals(link.getLabel(), "http://www.google.com");
		
		link = Ranges.range(sheet, "A2").getCellHyperlink();
		Assert.assertNotNull(link);
		Assert.assertEquals(Hyperlink.HyperlinkType.URL,link.getType());
		Assert.assertEquals(link.getAddress(), "http://www.zkoss.org");
		Assert.assertEquals(link.getLabel(), "http://www.zkoss.org");
		
		book = Util.swap(book);
		
		sheet = book.getSheetAt(0);
		link = Ranges.range(sheet, "A1").getCellHyperlink();
		Assert.assertNotNull(link);
		Assert.assertEquals(Hyperlink.HyperlinkType.URL,link.getType());
		Assert.assertEquals(link.getAddress(), "http://www.google.com");
		Assert.assertEquals(link.getLabel(), "http://www.google.com");
		
		link = Ranges.range(sheet, "A2").getCellHyperlink();
		Assert.assertNotNull(link);
		Assert.assertEquals(Hyperlink.HyperlinkType.URL,link.getType());
		Assert.assertEquals(link.getAddress(), "http://www.zkoss.org");
		Assert.assertEquals(link.getLabel(), "http://www.zkoss.org");
		
		Ranges.range(sheet, "A1").clearContents();
		
		Util.export(book, Setup.getTempFile());//get error if has this line
		
		link = Ranges.range(sheet, "A1").getCellHyperlink();
		Assert.assertNull(link);
		
		link = Ranges.range(sheet, "A2").getCellHyperlink();
		Assert.assertNotNull(link);
		Assert.assertEquals(Hyperlink.HyperlinkType.URL,link.getType());
		Assert.assertEquals(link.getAddress(), "http://www.zkoss.org");
		Assert.assertEquals(link.getLabel(), "http://www.zkoss.org");
		
	}

}
