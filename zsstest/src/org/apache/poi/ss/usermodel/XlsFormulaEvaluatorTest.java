/* FormulaEvaluatorTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 17, 2010 11:55:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.apache.poi.ss.usermodel;


import static org.junit.Assert.*;

import java.io.InputStream;

import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.usermodel.HSSFCell;
//import org.apache.poi.hssf.usermodel.HSSFEvaluationTestHelper;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;

/**
 * @author henrichen
 *
 */
public class XlsFormulaEvaluatorTest {
	private Workbook _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book1.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
		assertTrue(_workbook instanceof Book);
		assertTrue(_workbook instanceof HSSFBookImpl);
		assertTrue(_workbook instanceof HSSFWorkbook);
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
		assertNull(row.getCell(0));
		assertEquals(2, row.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, row.getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Cell cell = row.getCell(3); //D1: =SUM(A1:C1)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(5, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
		testToFormulaString(cell, "SUM(A1:C1)");
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
		testToFormulaString(cell, "SUM(Sheet2!A2:C2)");
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
		testToFormulaString(cell, "SUM('Sheet 1:Sheet3'!A3:C3)");
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
		
		testToFormulaString(cell, "SUM(Sheet2!A4)");
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
		
		testToFormulaString(cell, "SUM('Sheet 1:Sheet3'!A5)");
	}
	@Test
	public void testEvaluateAll() {
		testEvaluateArea();
		testEvaluateExternArea();
		testEvaluateArea3D();
		testEvaluateExternRef();
		testEvaluateRef3D();
	}
	
	private void testToFormulaString(Cell cell, String expect) {
		HSSFEvaluationWorkbook evalbook = HSSFEvaluationWorkbook.create((HSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
