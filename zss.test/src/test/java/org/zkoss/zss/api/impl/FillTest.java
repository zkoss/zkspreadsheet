package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.Range.AutoFillType;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

public class FillTest {
	
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
	public void testAutoFill2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testAutoFill(book);
	}
	
	@Test
	public void testAutoFill2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testAutoFill(book);
	}
	
	@Test
	public void testAutoFillMultiDim2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testAutoFillMultiDim(book);
	}
	
	@Test
	public void testAutoFillMultiDim2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testAutoFillMultiDim(book);
	}
	
	@Test
	public void testFillLeft2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testFillLeft(book);
	}
	
	@Test
	public void testFillLeft2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testFillLeft(book);
	}
	
	@Test
	public void testFillRight2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testFillRight(book);
	}
	
	@Test
	public void testFillRight2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testFillRight(book);
	}
	
	@Test
	public void testFillUp2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testFillUp(book);
	}
	
	@Test
	public void testFillUp2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testFillUp(book);
	}
	
	@Test
	public void testFillDown2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testFillDown(book);
	}
	
	@Test
	public void testFillDown2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testFillDown(book);
	}
	
	
	protected void testFillDown(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA1 = Ranges.range(sheet, "A1");
		rA1.setCellEditText("1");
		Range rA1A5 = Ranges.range(sheet, "A1:A5");
		rA1A5.fillDown();
		
		assertEquals("1", Ranges.range(sheet, "A2").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "A3").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "A4").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "A5").getCellEditText());
	}
	
	protected void testFillUp(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA5 = Ranges.range(sheet, "A5");
		rA5.setCellEditText("1");
		Range rA1A5 = Ranges.range(sheet, "A1:A5");
		rA1A5.fillUp();
		
		assertEquals("1", Ranges.range(sheet, "A1").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "A2").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "A3").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "A4").getCellEditText());
	}
	
	protected void testFillLeft(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rE1 = Ranges.range(sheet, "E1");
		rE1.setCellEditText("1");
		Range rA1E1 = Ranges.range(sheet, "A1:E1");
		rA1E1.fillLeft();
		
		assertEquals("1", Ranges.range(sheet, "A1").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "B1").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "C1").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "D1").getCellEditText());
	}
	
	protected void testFillRight(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA1 = Ranges.range(sheet, "A1");
		rA1.setCellEditText("1");
		Range rA1E1 = Ranges.range(sheet, "A1:E1");
		rA1E1.fillRight();
		
		assertEquals("1", Ranges.range(sheet, "B1").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "C1").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "D1").getCellEditText());
		assertEquals("1", Ranges.range(sheet, "E1").getCellEditText());
	}
	
	protected void testAutoFillMultiDim(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA1 = Ranges.range(sheet, "A1");
		rA1.setCellEditText("1");
		Range rA2 = Ranges.range(sheet, "A2");
		rA2.setCellEditText("2");
		Range rA3 = Ranges.range(sheet, "A3");
		rA3.setCellEditText("3");
		
		Range rB1 = Ranges.range(sheet, "B1");
		rB1.setCellEditText("4");
		Range rB2 = Ranges.range(sheet, "B2");
		rB2.setCellEditText("5");
		Range rB3 = Ranges.range(sheet, "B3");
		rB3.setCellEditText("6");
		
		Range rA1B3 = Ranges.range(sheet, "A1:B3");
		Range rangeA1D3 = Ranges.range(sheet, "A1:D3");
		rA1B3.autoFill(rangeA1D3, AutoFillType.DEFAULT);
		
		assertEquals("7", Ranges.range(sheet, "C1").getCellEditText());
		assertEquals("8", Ranges.range(sheet, "C2").getCellEditText());
		assertEquals("9", Ranges.range(sheet, "C3").getCellEditText());
		
		assertEquals("10", Ranges.range(sheet, "D1").getCellEditText());
		assertEquals("11", Ranges.range(sheet, "D2").getCellEditText());
		assertEquals("12", Ranges.range(sheet, "D3").getCellEditText());
		
	}
	
	protected void testAutoFill(Book workbook) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rA1 = Ranges.range(sheet, "A1");
		rA1.setCellEditText("1");
		Range rA2 = Ranges.range(sheet, "A2");
		rA2.setCellEditText("2");
		Range rA3 = Ranges.range(sheet, "A3");
		rA3.setCellEditText("3");
		Range rA1A3 = Ranges.range(sheet, "A1:A3");
		Range rangeA1A10 = Ranges.range(sheet, "A1:A10");
		rA1A3.autoFill(rangeA1A10, AutoFillType.DEFAULT);
		
		assertEquals("4", Ranges.range(sheet, "A4").getCellEditText());
		assertEquals("5", Ranges.range(sheet, "A5").getCellEditText());
		assertEquals("6", Ranges.range(sheet, "A6").getCellEditText());
		assertEquals("7", Ranges.range(sheet, "A7").getCellEditText());
		assertEquals("8", Ranges.range(sheet, "A8").getCellEditText());
		assertEquals("9", Ranges.range(sheet, "A9").getCellEditText());
		assertEquals("10", Ranges.range(sheet, "A10").getCellEditText());
	}
}
