/* NamedRangeXlsNamedRangeTest.java

	Purpose:
		
	Description:
		
	History:
		Nov 04, 2010 02:44:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.impl.ExcelImporter;
import org.zkoss.zss.model.sys.impl.HSSFBookImpl;

/**
 * @author henrichen
 *
 */
public class NamedRangeXlsNamedRangeTest {
	private XBook _workbook;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "NamedRange.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
		assertTrue(_workbook instanceof XBook);
		assertTrue(_workbook instanceof HSSFBookImpl);
		assertTrue(_workbook instanceof HSSFWorkbook);
		assertEquals(filename, ((XBook)_workbook).getBookName());
		assertEquals("Sheet1", _workbook.getSheetName(0));
		assertEquals("Sheet2", _workbook.getSheetName(1));
		assertEquals("Sheet3", _workbook.getSheetName(2));
		assertEquals(0, _workbook.getSheetIndex("Sheet1"));
		assertEquals(1, _workbook.getSheetIndex("Sheet2"));
		assertEquals(2, _workbook.getSheetIndex("Sheet3"));
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}

	@Test
	public void testNamedRange() {
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		
		XRange rngA1 = XRanges.range(sheet1, "RangeA1");
		assertEquals("Sheet1", rngA1.getSheet().getSheetName());
		assertEquals(0, rngA1.getColumn());
		assertEquals(0, rngA1.getRow());
		
		XRange rngB1 = XRanges.range(sheet1, "RangeB1");
		assertEquals("Sheet1", rngB1.getSheet().getSheetName());
		assertEquals(1, rngB1.getColumn());
		assertEquals(0, rngB1.getRow());
		
		XRange rngA2B3 = XRanges.range(sheet1, "RangeA2_B3");
		assertEquals("Sheet1", rngA2B3.getSheet().getSheetName());
		assertEquals(0, rngA2B3.getColumn());
		assertEquals(1, rngA2B3.getRow());
		assertEquals(1, rngA2B3.getLastColumn());
		assertEquals(2, rngA2B3.getLastRow());
	}
}
