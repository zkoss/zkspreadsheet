/* FormulaEvaluatorTest.java

	Purpose:
		
	Description:
		
	History:
		Mar 17, 2010 11:55:42 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;
import java.util.Set;

import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.usermodel.HSSFCell;
//xeimport org.apache.poi.hssf.usermodel.HSSFEvaluationTestHelper;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.impl.CellRefImpl;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;

/**
 * Test formula evaluation and hit test. 
 * @author henrichen
 *
 */
public class XlsFormulaHitTest {
	private Book _book;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book1.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_book = (Book) new ExcelImporter().imports(is, filename);
		assertTrue(_book instanceof Book);
		assertTrue(_book instanceof HSSFBookImpl);
		assertTrue(_book instanceof HSSFWorkbook);
		assertEquals(filename, ((Book)_book).getBookName());
		assertEquals("Sheet 1", _book.getSheetName(0));
		assertEquals("Sheet2", _book.getSheetName(1));
		assertEquals("Sheet3", _book.getSheetName(2));
		assertEquals(0, _book.getSheetIndex("Sheet 1"));
		assertEquals(1, _book.getSheetIndex("Sheet2"));
		assertEquals(2, _book.getSheetIndex("Sheet3"));

		_evaluator = ((Book)_book).getFormulaEvaluator();

	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_book = null;
		_evaluator = null;
	}

	@Test
	public void testEvaluateArea() {
		Sheet sheet1 = _book.getSheet("Sheet 1");
		Row row = sheet1.getRow(0);
		assertEquals(null, row.getCell(0));
		assertEquals(2, row.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, row.getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Cell cell = row.getCell(3); //D1: =SUM(A1:C1)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(5, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
		testToFormulaString(cell, "SUM(A1:C1)");
		
		Set<Ref>[] refs = new RangeImpl(0, 0, sheet1, sheet1).setValue(2); //A1 = 2
		Set<Ref> dependents = refs[0]; 
		assertEquals(1, dependents.size());
		final Ref ref = dependents.iterator().next();
		assertTrue(ref instanceof CellRefImpl);
		assertEquals(3, ref.getLeftCol());
		assertEquals(0, ref.getTopRow());
		
		BookHelper.clearFormulaCache(_book, refs[1]);
		value = _evaluator.evaluate(cell);
		assertEquals(7, value.getNumberValue(), 0.0000000000000001);
	}
	@Test
	public void testEvaluateExternArea() {
		Sheet sheet1 = _book.getSheet("Sheet 1");
		Sheet sheet2 = _book.getSheet("Sheet2");
		assertEquals(1, sheet2.getRow(1).getCell(0).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, sheet2.getRow(1).getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, sheet2.getRow(1).getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Row row = sheet1.getRow(1);
		Cell cell = row.getCell(3); //D2: =SUM(Sheet2!A2:C2)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(6, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
		testToFormulaString(cell, "SUM(Sheet2!A2:C2)");
		
		Set<Ref>[] refs = new RangeImpl(1, 0, sheet2, sheet2).setValue(2); //Sheet2!A2 = 2
		Set<Ref> dependents = refs[0]; 

		assertEquals(1, dependents.size());
		final Ref ref = dependents.iterator().next();
		assertTrue(ref instanceof CellRefImpl);
		assertEquals(3, ref.getLeftCol());
		assertEquals(1, ref.getTopRow());
		
		BookHelper.clearFormulaCache(_book, refs[1]);
		value = _evaluator.evaluate(cell);
		assertEquals(7, value.getNumberValue(), 0.0000000000000001);
	}
	@Test
	public void testEvaluateArea3D() {
		Sheet sheet1 = _book.getSheet("Sheet 1");
		Sheet sheet2 = _book.getSheet("Sheet2");
		Sheet sheet3 = _book.getSheet("Sheet3");
		assertEquals(1, sheet1.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, sheet2.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, sheet3.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Row row = sheet1.getRow(2);
		Cell cell = row.getCell(3); //D3: =SUM(Sheet1:Sheet3!A3:C3)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(6, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
		testToFormulaString(cell, "SUM('Sheet 1:Sheet3'!A3:C3)");

		Set<Ref>[] refs = new RangeImpl(2, 0, sheet1, sheet1).setValue(2); //Sheet1!A3 = 2
		Set<Ref> dependents = refs[0]; 

		assertEquals(1, dependents.size());
		final Ref ref = dependents.iterator().next();
		assertTrue(ref instanceof CellRefImpl);
		assertEquals(3, ref.getLeftCol());
		assertEquals(2, ref.getTopRow());
		
		BookHelper.clearFormulaCache(_book, refs[1]);
		value = _evaluator.evaluate(cell);
		assertEquals(8, value.getNumberValue(), 0.0000000000000001);
	}
	@Test
	public void testEvaluateExternRef() {
		Sheet sheet1 = _book.getSheet("Sheet 1");
		Sheet sheet2 = _book.getSheet("Sheet2");
		assertEquals(1, sheet2.getRow(1).getCell(0).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, sheet2.getRow(1).getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, sheet2.getRow(1).getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Row row = sheet1.getRow(3);
		Cell cell = row.getCell(3); //D4: =SUM(Sheet2!A4)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(1, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
		testToFormulaString(cell, "SUM(Sheet2!A4)");

		Set<Ref>[] refs = new RangeImpl(3, 0, sheet2, sheet2).setValue(2); //Sheet2!A4 = 2
		Set<Ref> dependents = refs[0]; 

		assertEquals(1, dependents.size());
		final Ref ref = dependents.iterator().next();
		assertTrue(ref instanceof CellRefImpl);
		assertEquals(3, ref.getLeftCol());
		assertEquals(3, ref.getTopRow());
		
		BookHelper.clearFormulaCache(_book, refs[1]);
		value = _evaluator.evaluate(cell);
		assertEquals(2, value.getNumberValue(), 0.0000000000000001);
	}
	@Test
	public void testEvaluateRef3D() {
		Sheet sheet1 = _book.getSheet("Sheet 1");
		Sheet sheet2 = _book.getSheet("Sheet2");
		Sheet sheet3 = _book.getSheet("Sheet3");
		assertEquals(1, sheet1.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, sheet2.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, sheet3.getRow(2).getCell(2).getNumericCellValue(), 0.0000000000000001);
		
		Row row = sheet1.getRow(4);
		Cell cell = row.getCell(3); //D5: =SUM(Sheet1:Sheet3!A5)
		CellValue value = _evaluator.evaluate(cell);
		assertEquals(6, value.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, value.getCellType());
		
		testToFormulaString(cell, "SUM('Sheet 1:Sheet3'!A5)");
		
		Set<Ref>[] refs = new RangeImpl(4, 0, sheet1, sheet1).setValue(2); //Sheet1!A5 = 2
		Set<Ref> dependents = refs[0]; 

		assertEquals(1, dependents.size());
		final Ref ref = dependents.iterator().next();
		assertTrue(ref instanceof CellRefImpl);
		assertEquals(3, ref.getLeftCol());
		assertEquals(4, ref.getTopRow());
		
		BookHelper.clearFormulaCache(_book, refs[1]);
		value = _evaluator.evaluate(cell);
		assertEquals(7, value.getNumberValue(), 0.0000000000000001);
		
	}
	
	private void testToFormulaString(Cell cell, String expect) {
/*		EvaluationCell srcCell = HSSFEvaluationTestHelper.wrapCell((HSSFCell)cell);
		HSSFEvaluationWorkbook evalbook = HSSFEvaluationWorkbook.create((HSSFWorkbook)_book);
		Ptg[] ptgs = evalbook.getFormulaTokens(srcCell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
*/	}
}
