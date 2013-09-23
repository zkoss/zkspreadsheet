package org.zkoss.zss.api.impl;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;

public class RangeAPITest extends RangeAPITestBase {
	
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
	public void testCreateSheet2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testCreateSheet(book);
	}
	
	@Test
	public void testCreateSheet2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testCreateSheet(book);
	}
	
	@Test
	public void testDeleteSheet2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testDeleteSheet(book);
	}
	
	@Test
	public void testDeleteSheet2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testDeleteSheet(book);
	}
	
	@Test
	public void testGetColumn2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testGetColumn(book);
	}
	
	@Test
	public void testGetColumn2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testGetColumn(book);
	}
	
	@Test
	public void testGetRow2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testGetRow(book);
	}
	
	@Test
	public void testGetRow2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testGetRow(book);
	}
	
	@Test
	public void testGetLastColumn2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testGetLastColumn(book);
	}
	
	@Test
	public void testGetLastColumn2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testGetLastColumn(book);
	}
	
	@Test
	public void testGetLastRow2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testGetLastRow(book);
	}
	
	@Test
	public void testGetLastRow2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testGetLastRow(book);
	}
	
	@Test
	public void testGetRowCount2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testGetRowCount(book);
	}
	
	@Test
	public void testGetRowCount2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testGetRowCount(book);
	}
	
	@Test
	public void testGetColumnCount2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testGetColumnCount(book);
	}
	
	@Test
	public void testGetColumnCount2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testGetColumnCount(book);
	}
	
	@Test
	public void testIsWholeRow2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testIsWholeRow(book);
	}
	
	@Test
	public void testIsWholeRow2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testIsWholeRow(book);
	}
	
	@Test
	public void testIsWholeColumn2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testIsWholeColumn(book);
	}
	
	@Test
	public void testIsWholeColumn2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testIsWholeColumn(book);
	}
	
	@Test
	public void testIsWholeSheet2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testIsWholeSheet(book);
	}
	
	@Test
	public void testIsWholeSheet2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testIsWholeSheet(book);
	}
	
	@Test
	public void testClearContents2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testClearContents(book);
	}
	
	@Test
	public void testClearContents2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testIsWholeColumn(book);
	}
	
	@Test
	public void testClearStyle2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testClearStyle(book);
	}
	
	@Test
	public void testClearStyle2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testClearStyle(book);
	}
	
	@Test
	public void testHyperLink2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testHyperLink(book, Setup.getTempFile());
	}
	
	@Test
	public void testHyperLink2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testHyperLink(book, Setup.getTempFile());
	}
	
	@Test
	public void testToShiftRanged2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testToShiftRanged(book);
	}
	
	@Test
	public void testToShiftRanged2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testToShiftRanged(book);
	}
	
	@Test
	public void testToCellRange2003() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xls");
		testToCellRange(book);
	}
	
	@Test
	public void testToCellRange2007() throws IOException {
		Book book = Util.loadBook(this,"book/blank.xlsx");
		testToCellRange(book);
	}
	
}
