/* FormulaEvaluatorTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 17, 2010 11:55:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.apache.poi.ss.usermodel;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BooksImpl;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;
import org.zkoss.zss.model.impl.XSSFBookImpl;

/**
 * @author henrichen
 *
 */
public class Book1XlsExternalBookReferenceTest {
	private Workbook _workbook1;
	private Workbook _workbook2;
	private BooksImpl _books;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename1 = "Book1.xls";
		final InputStream is1 = new ClassLocator().getResourceAsStream(filename1);
		_workbook1 = new ExcelImporter().imports(is1, filename1);
		assertTrue(_workbook1 instanceof Book);
		assertTrue(_workbook1 instanceof HSSFBookImpl);
		assertTrue(_workbook1 instanceof HSSFWorkbook);
		assertEquals(filename1, ((Book)_workbook1).getBookName());
		assertEquals("Sheet 1", _workbook1.getSheetName(0));
		assertEquals("Sheet2", _workbook1.getSheetName(1));
		assertEquals("Sheet3", _workbook1.getSheetName(2));
		assertEquals(0, _workbook1.getSheetIndex("Sheet 1"));
		assertEquals(1, _workbook1.getSheetIndex("Sheet2"));
		assertEquals(2, _workbook1.getSheetIndex("Sheet3"));
		
		final String filename2 = "Book1.xlsx";
		final InputStream is2 = new ClassLocator().getResourceAsStream(filename2);
		_workbook2 = new ExcelImporter().imports(is2, filename2);
		assertTrue(_workbook2 instanceof Book);
		assertTrue(_workbook2 instanceof XSSFBookImpl);
		assertTrue(_workbook2 instanceof XSSFWorkbook);
		assertEquals(filename2, ((Book)_workbook2).getBookName());
		assertEquals("Sheet 1", _workbook1.getSheetName(0));
		assertEquals("Sheet2", _workbook1.getSheetName(1));
		assertEquals("Sheet3", _workbook1.getSheetName(2));
		assertEquals(0, _workbook2.getSheetIndex("Sheet 1"));
		assertEquals(1, _workbook2.getSheetIndex("Sheet2"));
		assertEquals(2, _workbook2.getSheetIndex("Sheet3"));

		_books = new BooksImpl(new Book[] {(Book)_workbook1, (Book)_workbook2});
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
		Cell cell = row.getCell(3); //D6: =SUM('[Book1.xls]Sheet 1'!C1)
		CellValue value = ((Book)_workbook1).getFormulaEvaluator().evaluate(cell);
		assertEquals(3, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
	}
}
