package org.zkoss.zss.api.impl;

import static org.junit.Assert.assertTrue;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.ZssContext;
import org.zkoss.zss.Setup;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;

public class MergeTest {
	
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
	
	protected void testUnMerge(Book workbook, String outFileName) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertTrue(!Util.isAMergedRange(rangeA5F18));
		rangeA5F18.merge(false);
		assertTrue(Util.isAMergedRange(rangeA5F18));
		rangeA5F18.unmerge();
		assertTrue(!Util.isAMergedRange(rangeA5F18));
		
		Util.export(workbook, outFileName);
		
		assertTrue(!Util.isAMergedRange(rangeA5F18));
		
		workbook = Util.loadBook(outFileName);
		sheet = workbook.getSheet("Sheet1");
		rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertTrue(!Util.isAMergedRange(rangeA5F18));
		
	}
	
	// Also check the cell is merged after export
	protected void testMerge(Book workbook, String outFileName) throws IOException {
		Sheet sheet = workbook.getSheet("Sheet1");
		Range rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertTrue(!Util.isAMergedRange(rangeA5F18));
		rangeA5F18.merge(false);
		assertTrue(Util.isAMergedRange(rangeA5F18));
		
		Util.export(workbook, outFileName);
		
		assertTrue(Util.isAMergedRange(rangeA5F18));
		
		workbook = Util.loadBook(outFileName);
		sheet = workbook.getSheet("Sheet1");
		rangeA5F18 = Ranges.range(sheet, "A5:F18");
		assertTrue(Util.isAMergedRange(rangeA5F18));
	}
}
