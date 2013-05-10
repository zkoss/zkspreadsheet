/* FormulaEvaluatorTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 17, 2010 11:55:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.poi.ss.usermodel;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XBookSeries;
import org.zkoss.zss.model.sys.impl.ExcelImporter;
import org.zkoss.zss.model.sys.impl.HSSFBookImpl;
import org.zkoss.zss.model.sys.impl.XSSFBookImpl;
import org.zkoss.zssex.model.impl.BookSeriesImpl;

/**
 * @author henrichen
 *
 */
public class Book1XlsxExternalBookReferenceTest {
	private Workbook _workbook1;
	private Workbook _workbook2;
	private XBookSeries _books;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename1 = "Book1.xlsx";
		final InputStream is1 = new ClassLocator().getResourceAsStream(filename1);
		_workbook1 = new ExcelImporter().imports(is1, filename1);
		assertTrue(_workbook1 instanceof XBook);
		assertTrue(_workbook1 instanceof XSSFBookImpl);
		assertTrue(_workbook1 instanceof XSSFWorkbook);
		assertEquals(filename1, ((XBook)_workbook1).getBookName());
		assertEquals("Sheet 1", _workbook1.getSheetName(0));
		assertEquals("Sheet2", _workbook1.getSheetName(1));
		assertEquals("Sheet3", _workbook1.getSheetName(2));
		assertEquals(0, _workbook1.getSheetIndex("Sheet 1"));
		assertEquals(1, _workbook1.getSheetIndex("Sheet2"));
		assertEquals(2, _workbook1.getSheetIndex("Sheet3"));
		
		final String filename2 = "Book1.xls";
		final InputStream is2 = new ClassLocator().getResourceAsStream(filename2);
		_workbook2 = new ExcelImporter().imports(is2, filename2);
		assertTrue(_workbook2 instanceof XBook);
		assertTrue(_workbook2 instanceof HSSFBookImpl);
		assertTrue(_workbook2 instanceof HSSFWorkbook);
		assertEquals(filename2, ((XBook)_workbook2).getBookName());
		assertEquals("Sheet 1", _workbook1.getSheetName(0));
		assertEquals("Sheet2", _workbook1.getSheetName(1));
		assertEquals("Sheet3", _workbook1.getSheetName(2));
		assertEquals(0, _workbook2.getSheetIndex("Sheet 1"));
		assertEquals(1, _workbook2.getSheetIndex("Sheet2"));
		assertEquals(2, _workbook2.getSheetIndex("Sheet3"));
		
		_books = new BookSeriesImpl(new XBook[] {(XBook) _workbook1, (XBook) _workbook2});
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_workbook1 = null;
		_workbook2 = null;
		_books = null;
	}

	@Test
	public void testExternalReferenceRid() {
		Sheet sheet1 = _workbook1.getSheet("Sheet 1");
		Row row = sheet1.getRow(5);
		Cell cell = row.getCell(3); //D6: =SUM('[Book1.xls]Sheet 1'!A1:C1)
		CellValue value = ((XBook)_workbook1).getFormulaEvaluator().evaluate(cell);
		assertEquals(5, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
	}
}
