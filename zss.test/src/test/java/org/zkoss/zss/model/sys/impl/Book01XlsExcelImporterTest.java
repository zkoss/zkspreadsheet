/* ExcelImporterTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 15, 2010 3:31:34 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys.impl;


import java.io.IOException;
import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;

import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.ExcelImporter;
import org.zkoss.zss.model.sys.impl.HSSFBookImpl;
import org.zkoss.zss.model.sys.impl.XSSFBookImpl;

/**
 * @author henrichen
 *
 */
public class Book01XlsExcelImporterTest {

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
	}

	@Test
	public void testImportsFromXlsInputStream() throws IOException{
		final String filename = "Book1.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		Workbook wb = new ExcelImporter().imports(is, filename);
		assertTrue(wb instanceof XBook);
		assertTrue(wb instanceof HSSFBookImpl);
		assertTrue(wb instanceof HSSFWorkbook);
		assertEquals(filename, ((XBook)wb).getBookName());
		assertEquals("Sheet 1", wb.getSheetName(0));
		assertEquals("Sheet2", wb.getSheetName(1));
		assertEquals("Sheet3", wb.getSheetName(2));
		assertEquals(0, wb.getSheetIndex("Sheet 1"));
		assertEquals(1, wb.getSheetIndex("Sheet2"));
		assertEquals(2, wb.getSheetIndex("Sheet3"));
	}
	
	@Test
	public void testImportsFromXlsxInputStream() throws IOException{
		final String filename = "Book1.xlsx";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		Workbook wb = new ExcelImporter().imports(is, filename);
		assertTrue(wb instanceof XBook);
		assertTrue(wb instanceof XSSFBookImpl);
		assertTrue(wb instanceof XSSFWorkbook);
		assertEquals(filename, ((XBook)wb).getBookName());
		assertEquals("Sheet 1", wb.getSheetName(0));
		assertEquals("Sheet2", wb.getSheetName(1));
		assertEquals("Sheet3", wb.getSheetName(2));
		assertEquals(0, wb.getSheetIndex("Sheet 1"));
		assertEquals(1, wb.getSheetIndex("Sheet2"));
		assertEquals(2, wb.getSheetIndex("Sheet3"));
	}
}
