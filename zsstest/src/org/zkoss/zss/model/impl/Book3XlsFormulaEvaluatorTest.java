/* FormulaEvaluatorTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 17, 2010 11:55:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;


import static org.junit.Assert.*;

import java.io.InputStream;

import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFEvaluationTestHelper;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;

/**
 * @author henrichen
 *
 */
public class Book3XlsFormulaEvaluatorTest {
	private Book _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book3.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
		assertTrue(_workbook instanceof Book);
		assertTrue(_workbook instanceof HSSFBookImpl);
		assertTrue(_workbook instanceof HSSFWorkbook);
		assertEquals(filename, ((Book)_workbook).getBookName());
		assertEquals("Sheet1", _workbook.getSheetName(0));
		assertEquals("Sheet2", _workbook.getSheetName(1));
		assertEquals("Sheet3", _workbook.getSheetName(2));
		assertEquals(0, _workbook.getSheetIndex("Sheet1"));
		assertEquals(1, _workbook.getSheetIndex("Sheet2"));
		assertEquals(2, _workbook.getSheetIndex("Sheet3"));

		_evaluator = ((Book)_workbook).getFormulaEvaluator();
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_workbook = null;
		_evaluator = null;
	}

	@Test
	public void testEvaluateArea() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Row row = sheet1.getRow(0);
		assertNull(row.getCell(0));
		
		
		for(int col = 1; col < 13; ++col) {
			final Cell cell = row.getCell(col);
			CellValue value = _evaluator.evaluate(cell);
			assertEquals(0, value.getNumberValue(), 0.0000000000000001);
		}

		Row row4 = sheet1.getRow(3);
		final Cell cell = row4.getCell(0);
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(0, value.getNumberValue(), 0.0000000000000001);
		
		Cell cellA1 = row.getCell(0);
		RichTextString rstr = _workbook.getCreationHelper().createRichTextString("10");
		new RangeImpl(0, 0, sheet1, sheet1).setRichEditText(rstr);
	}
}
