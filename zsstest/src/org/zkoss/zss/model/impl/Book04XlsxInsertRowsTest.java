/**
 * 
 */
package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.hssf.record.formula.Ptg;
import org.zkoss.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.ss.usermodel.FormulaEvaluator;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;

/**
 * Insert a row and check if the formula still work
 * @author henrichen
 */
public class Book04XlsxInsertRowsTest {
	private Book _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book4.xlsx";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
		assertTrue(_workbook instanceof Book);
		assertTrue(_workbook instanceof XSSFBookImpl);
		assertTrue(_workbook instanceof XSSFWorkbook);
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
	public void testInsertA5_D5() {
		Worksheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row8 = sheet1.getRow(7);
		assertEquals(1, row8.getCell(0).getNumericCellValue(), 0.0000000000000001); //A8: 1
		assertEquals(2, row8.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B8: 2
		assertEquals(3, row8.getCell(2).getNumericCellValue(), 0.0000000000000001); //C8: 3
		assertEquals(7, row1.getCell(3).getNumericCellValue(), 0.0000000000000001); //D1: 7
		
		//A1: =SUM(A8:C8)
		Cell cellA1 = row1.getCell(0); //A1: =SUM(A8:C8)
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(6, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(A8:C8)");
		
		//B1: =SUM(B5:B1048576)
		Cell cellB1 = row1.getCell(1);
		CellValue valueB1 = _evaluator.evaluate(cellB1);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueB1.getCellType());
		testToFormulaString(cellB1, "SUM(B5:B1048576)");
		assertEquals(2, valueB1.getNumberValue(), 0.0000000000000001);

		//C1: =SUM(B1048575:B1048576)
		Cell cellC1 = row1.getCell(2);
		CellValue valueC1 = _evaluator.evaluate(cellC1);
		assertEquals(0, valueC1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC1.getCellType());
		testToFormulaString(cellC1, "SUM(B1048575:B1048576)");
		
		//A10: =SUM(A8:C8)
		Row row10 = sheet1.getRow(9);
		Cell cellA10 = row10.getCell(0); 
		CellValue valueA10 = _evaluator.evaluate(cellA10);
		assertEquals(6, valueA10.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA10.getCellType());
		testToFormulaString(cellA10, "SUM(A8:C8)");

		//D10: =D1
		Cell cellD10 = row10.getCell(3); 
		CellValue valueD10 = _evaluator.evaluate(cellD10);
		assertEquals(7, valueD10.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueD10.getCellType());
		testToFormulaString(cellD10, "D1");
		
		//Insert A5:D5
		BookHelper.insertRange(sheet1, 4, 0, 4, 3, false, Range.FORMAT_LEFTABOVE);
		_evaluator.notifySetFormula(cellC1);
		
		Row row9 = sheet1.getRow(8);
		assertEquals(1, row9.getCell(0).getNumericCellValue(), 0.0000000000000001); //A9: 1
		assertEquals(2, row9.getCell(1).getNumericCellValue(), 0.0000000000000001); //B9: 2
		assertEquals(3, row9.getCell(2).getNumericCellValue(), 0.0000000000000001); //C9: 3
		assertEquals(7, row1.getCell(3).getNumericCellValue(), 0.0000000000000001); //D1: 7

		//A1: =SUM(A9:C9)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(6, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(A9:C9)");
		
		//B1: =SUM(B6:B1048576)
		cellB1 = row1.getCell(1);
		valueB1 = _evaluator.evaluate(cellB1);
		assertEquals(2, valueB1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueB1.getCellType());
		testToFormulaString(cellB1, "SUM(B6:B1048576)");

		//C1: =SUM(B1048576:B1048576)
		cellC1 = row1.getCell(2);
		valueC1 = _evaluator.evaluate(cellC1);
		assertEquals(0, valueC1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC1.getCellType());
		testToFormulaString(cellC1, "SUM(B1048576:B1048576)");
		
		//A11: =SUM(A9:C9)
		Row row11 = sheet1.getRow(10);
		Cell cellA11 = row11.getCell(0); 
		CellValue valueA11 = _evaluator.evaluate(cellA11);
		assertEquals(6, valueA11.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA11.getCellType());
		testToFormulaString(cellA11, "SUM(A9:C9)");

		//D11: =D1
		Cell cellD11 = row11.getCell(3); //D11: =D1
		CellValue valueD11 = _evaluator.evaluate(cellD11);
		assertEquals(7, valueD11.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueD11.getCellType());
		testToFormulaString(cellD11, "D1");

	}
	
	@Test
	public void testInsertRow5() {
		Worksheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row8 = sheet1.getRow(7);
		assertEquals(1, row8.getCell(0).getNumericCellValue(), 0.0000000000000001); //A8: 1
		assertEquals(2, row8.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B8: 2
		assertEquals(3, row8.getCell(2).getNumericCellValue(), 0.0000000000000001); //C8: 3
		assertEquals(7, row1.getCell(3).getNumericCellValue(), 0.0000000000000001); //D1: 7
		
		//A1: =SUM(A8:C8)
		Cell cellA1 = row1.getCell(0); //A1: =SUM(A8:C8)
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(6, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(A8:C8)");
		
		//B1: =SUM(B5:B1048576)
		Cell cellB1 = row1.getCell(1);
		CellValue valueB1 = _evaluator.evaluate(cellB1);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueB1.getCellType());
		testToFormulaString(cellB1, "SUM(B5:B1048576)");
		assertEquals(2, valueB1.getNumberValue(), 0.0000000000000001);

		//C1: =SUM(B1048575:B1048576)
		Cell cellC1 = row1.getCell(2);
		CellValue valueC1 = _evaluator.evaluate(cellC1);
		assertEquals(0, valueC1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC1.getCellType());
		testToFormulaString(cellC1, "SUM(B1048575:B1048576)");
		
		//A10: =SUM(A8:C8)
		Row row10 = sheet1.getRow(9);
		Cell cellA10 = row10.getCell(0); 
		CellValue valueA10 = _evaluator.evaluate(cellA10);
		assertEquals(6, valueA10.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA10.getCellType());
		testToFormulaString(cellA10, "SUM(A8:C8)");

		//D10: =D1
		Cell cellD10 = row10.getCell(3); 
		CellValue valueD10 = _evaluator.evaluate(cellD10);
		assertEquals(7, valueD10.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueD10.getCellType());
		testToFormulaString(cellD10, "D1");
		
		//Insert row 5
		BookHelper.insertRows(sheet1, 4, 1, Range.FORMAT_LEFTABOVE);
		_evaluator.notifySetFormula(cellC1);
		
		//height shall be the same
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		assertEquals(row4.getHeight(), row5.getHeight());
		
		Row row9 = sheet1.getRow(8);
		assertEquals(1, row9.getCell(0).getNumericCellValue(), 0.0000000000000001); //A9: 1
		assertEquals(2, row9.getCell(1).getNumericCellValue(), 0.0000000000000001); //B9: 2
		assertEquals(3, row9.getCell(2).getNumericCellValue(), 0.0000000000000001); //C9: 3
		assertEquals(7, row1.getCell(3).getNumericCellValue(), 0.0000000000000001); //D1: 7

		//A1: =SUM(A9:C9)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(6, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(A9:C9)");
		
		//B1: =SUM(B6:B1048576)
		cellB1 = row1.getCell(1);
		valueB1 = _evaluator.evaluate(cellB1);
		assertEquals(2, valueB1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueB1.getCellType());
		testToFormulaString(cellB1, "SUM(B6:B1048576)");

		//C1: =SUM(B1048576:B1048576)
		cellC1 = row1.getCell(2);
		valueC1 = _evaluator.evaluate(cellC1);
		assertEquals(0, valueC1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC1.getCellType());
		testToFormulaString(cellC1, "SUM(B1048576:B1048576)");
		
		//A11: =SUM(A9:C9)
		Row row11 = sheet1.getRow(10);
		Cell cellA11 = row11.getCell(0); 
		CellValue valueA11 = _evaluator.evaluate(cellA11);
		assertEquals(6, valueA11.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA11.getCellType());
		testToFormulaString(cellA11, "SUM(A9:C9)");

		//D11: =D1
		Cell cellD11 = row11.getCell(3); //D11: =D1
		CellValue valueD11 = _evaluator.evaluate(cellD11);
		assertEquals(7, valueD11.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueD11.getCellType());
		testToFormulaString(cellD11, "D1");

	}
	
	private void testToFormulaString(Cell cell, String expect) {
		XSSFEvaluationWorkbook evalbook = XSSFEvaluationWorkbook.create((XSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
