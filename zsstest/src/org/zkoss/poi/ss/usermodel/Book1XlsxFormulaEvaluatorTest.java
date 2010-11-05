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
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.XSSFBookImpl;

/**
 * @author henrichen
 *
 */
public class Book1XlsxFormulaEvaluatorTest {
	private Workbook _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book1.xlsx";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
		assertTrue(_workbook instanceof Book);
		assertTrue(_workbook instanceof XSSFBookImpl);
		assertTrue(_workbook instanceof XSSFWorkbook);
		assertEquals(filename, ((Book)_workbook).getBookName());
		assertEquals("Sheet 1", _workbook.getSheetName(0));
		assertEquals("Sheet2", _workbook.getSheetName(1));
		assertEquals("Sheet3", _workbook.getSheetName(2));
		assertEquals(0, _workbook.getSheetIndex("Sheet 1"));
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
		Sheet sheet1 = _workbook.getSheet("Sheet 1");
		Row row = sheet1.getRow(0);
		assertEquals(1, row.getCell(0).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, row.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, row.getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Cell cell = row.getCell(3); //D1: =SUM(A1:C1)
		_evaluator.evaluate(cell);
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(6, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
	}
	@Test
	public void testEvaluateExternArea() {
		Sheet sheet1 = _workbook.getSheet("Sheet 1");
		Sheet sheet2 = _workbook.getSheet("Sheet2");
		assertEquals(1, sheet2.getRow(1).getCell(0).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, sheet2.getRow(1).getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, sheet2.getRow(1).getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Row row = sheet1.getRow(1);
		Cell cell = row.getCell(3); //D2: =SUM(Sheet2!A2:C2)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(6, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
	}
	@Test
	public void testEvaluateArea3D() {
		Sheet sheet1 = _workbook.getSheet("Sheet 1");
		Sheet sheet2 = _workbook.getSheet("Sheet2");
		Sheet sheet3 = _workbook.getSheet("Sheet3");
		assertEquals(1, sheet1.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, sheet2.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, sheet3.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Row row = sheet1.getRow(2);
		Cell cell = row.getCell(3); //D3: =SUM(Sheet1:Sheet3!A3:C3)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(6, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
	}
	@Test
	public void testEvaluateExternRef() {
		Sheet sheet1 = _workbook.getSheet("Sheet 1");
		Sheet sheet2 = _workbook.getSheet("Sheet2");
		assertEquals(1, sheet2.getRow(1).getCell(0).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, sheet2.getRow(1).getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, sheet2.getRow(1).getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Row row = sheet1.getRow(3);
		Cell cell = row.getCell(3); //D4: =SUM(Sheet2!A4)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(1, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
	}
	@Test
	public void testEvaluateRef3D() {
		Sheet sheet1 = _workbook.getSheet("Sheet 1");
		Sheet sheet2 = _workbook.getSheet("Sheet2");
		Sheet sheet3 = _workbook.getSheet("Sheet3");
		assertEquals(1, sheet1.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, sheet2.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, sheet3.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Row row = sheet1.getRow(4);
		Cell cell = row.getCell(3); //D5: =SUM(Sheet1:Sheet3!A5)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(6, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
	}
	@Test
	public void testEvaluateAll() {
		testEvaluateArea();
		testEvaluateExternArea();
		testEvaluateArea3D();
		testEvaluateExternRef();
		testEvaluateRef3D();
	}
}
