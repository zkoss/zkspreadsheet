/* ExcelImporterTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 15, 2010 3:31:34 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;


import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import static org.junit.Assert.*;

import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;

/**
 * @author henrichen
 *
 */
public class ExcelImporterTest {

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
	public void testImportsFromXlsInputStream() {
		final String filename = "Book1.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		Workbook wb = new ExcelImporter().imports(is, filename);
		assertTrue(wb instanceof Book);
		assertTrue(wb instanceof HSSFBookImpl);
		assertTrue(wb instanceof HSSFWorkbook);
		assertEquals(filename, ((Book)wb).getBookName());
		assertEquals("Sheet 1", wb.getSheetName(0));
		assertEquals("Sheet2", wb.getSheetName(1));
		assertEquals("Sheet3", wb.getSheetName(2));
		assertEquals(0, wb.getSheetIndex("Sheet 1"));
		assertEquals(1, wb.getSheetIndex("Sheet2"));
		assertEquals(2, wb.getSheetIndex("Sheet3"));
	}
	
	@Test
	public void testImportsFromXlsxInputStream() {
		final String filename = "Book1.xlsx";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		Workbook wb = new ExcelImporter().imports(is, filename);
		assertTrue(wb instanceof Book);
		assertTrue(wb instanceof XSSFBookImpl);
		assertTrue(wb instanceof XSSFWorkbook);
		assertEquals(filename, ((Book)wb).getBookName());
		assertEquals("Sheet 1", wb.getSheetName(0));
		assertEquals("Sheet2", wb.getSheetName(1));
		assertEquals("Sheet3", wb.getSheetName(2));
		assertEquals(0, wb.getSheetIndex("Sheet 1"));
		assertEquals(1, wb.getSheetIndex("Sheet2"));
		assertEquals(2, wb.getSheetIndex("Sheet3"));
	}
}
