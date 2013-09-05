package org.zkoss.zss.api.impl;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.model.Book;

public class RangeAPITest extends RangeAPITestBase {
	
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
	public void testCreateSheet2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testCreateSheet(book);
	}
	
	@Test
	public void testCreateSheet2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testCreateSheet(book);
	}
	
	@Test
	public void testDeleteSheet2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testDeleteSheet(book);
	}
	
	@Test
	public void testDeleteSheet2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testDeleteSheet(book);
	}
	
	@Test
	public void testUnMerge2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testUnMerge(book, "book/test.xls");
	}
	
	@Test
	public void testUnMerge2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testUnMerge(book, "book/test.xlsx");
	}
	
	@Test
	public void testMerge2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testMerge(book, "book/test.xls");
	}
	
	@Test
	public void testMerge2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testMerge(book, "book/test.xlsx");
	}
	
	@Test
	public void testGetColumn2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testGetColumn(book);
	}
	
	@Test
	public void testGetColumn2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testGetColumn(book);
	}
	
	@Test
	public void testGetRow2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testGetRow(book);
	}
	
	@Test
	public void testGetRow2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testGetRow(book);
	}
	
	@Test
	public void testGetLastColumn2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testGetLastColumn(book);
	}
	
	@Test
	public void testGetLastColumn2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testGetLastColumn(book);
	}
	
	@Test
	public void testGetLastRow2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testGetLastRow(book);
	}
	
	@Test
	public void testGetLastRow2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testGetLastRow(book);
	}
	
	@Test
	public void testGetRowCount2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testGetRowCount(book);
	}
	
	@Test
	public void testGetRowCount2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testGetRowCount(book);
	}
	
	@Test
	public void testGetColumnCount2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testGetColumnCount(book);
	}
	
	@Test
	public void testGetColumnCount2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testGetColumnCount(book);
	}
	
	@Test
	public void testIsWholeRow2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testIsWholeRow(book);
	}
	
	@Test
	public void testIsWholeRow2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testIsWholeRow(book);
	}
	
	@Test
	public void testIsWholeColumn2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testIsWholeColumn(book);
	}
	
	@Test
	public void testIsWholeColumn2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testIsWholeColumn(book);
	}
	
	@Test
	public void testIsWholeSheet2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testIsWholeSheet(book);
	}
	
	@Test
	public void testIsWholeSheet2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testIsWholeSheet(book);
	}
	
	@Test
	public void testClearContents2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testClearContents(book);
	}
	
	@Test
	public void testClearContents2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testIsWholeColumn(book);
	}
	
	@Test
	public void testClearStyle2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testClearStyle(book);
	}
	
	@Test
	public void testClearStyle2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testClearStyle(book);
	}
	
	@Test
	public void testHyperLink2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testHyperLink(book, "book/test.xls");
	}
	
	@Test
	public void testHyperLink2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testHyperLink(book, "book/test.xlsx");
	}
	
	@Test
	public void testAutoFill2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testAutoFill(book);
	}
	
	@Test
	public void testAutoFill2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testAutoFill(book);
	}
	
	@Test
	public void testAutoFillMultiDim2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testAutoFillMultiDim(book);
	}
	
	@Test
	public void testAutoFillMultiDim2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testAutoFillMultiDim(book);
	}
	
	@Test
	public void testFillLeft2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testFillLeft(book);
	}
	
	@Test
	public void testFillLeft2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testFillLeft(book);
	}
	
	@Test
	public void testFillRight2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testFillRight(book);
	}
	
	@Test
	public void testFillRight2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testFillRight(book);
	}
	
	@Test
	public void testFillUp2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testFillUp(book);
	}
	
	@Test
	public void testFillUp2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testFillUp(book);
	}
	
	@Test
	public void testFillDown2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testFillDown(book);
	}
	
	@Test
	public void testFillDown2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testFillDown(book);
	}
	
	@Test
	public void testToShiftRanged2003() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testToShiftRanged(book);
	}
	
	@Test
	public void testToShiftRanged2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testToShiftRanged(book);
	}
	
	@Test
	public void testToCellRange2003() throws IOException {
		Book book = Util.loadBook("book/blank.xls");
		testToCellRange(book);
	}
	
	@Test
	public void testToCellRange2007() throws IOException {
		Book book = Util.loadBook("book/blank.xlsx");
		testToCellRange(book);
	}
	
}
