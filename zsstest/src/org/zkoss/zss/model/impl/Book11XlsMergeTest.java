/**
 * 
 */
package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;
import java.util.Set;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.hssf.record.constant.ErrorConstant;
import org.zkoss.poi.hssf.record.formula.Ptg;
import org.zkoss.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.ss.usermodel.ErrorConstants;
import org.zkoss.poi.ss.usermodel.FormulaEvaluator;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.impl.CellRefImpl;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;
import org.zkoss.zss.ui.impl.Utils;

/**
 * Insert a row and check if the formula still work
 * @author henrichen
 */
public class Book11XlsMergeTest {
	private Book _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book11.xls";
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
	public void testMergeCells() {
		Worksheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row2 = sheet1.getRow(1);
		
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		Row row9 = sheet1.getRow(8);
		Row row10 = sheet1.getRow(9);
		Row row11 = sheet1.getRow(10);
		Row row12 = sheet1.getRow(11);

		//merged range (A1:B2)
		assertEquals(Cell.CELL_TYPE_BLANK, row1.getCell(0).getCellType()); //A1
		assertEquals(Cell.CELL_TYPE_BLANK, row1.getCell(1).getCellType()); //A2
		assertNull(row1.getCell(2)); //A3
		assertEquals(Cell.CELL_TYPE_BLANK, row2.getCell(0).getCellType()); //B1
		assertEquals(Cell.CELL_TYPE_BLANK, row2.getCell(1).getCellType()); //B2
		assertNull(row2.getCell(2)); //B3
		assertEquals(1, row9.getCell(0).getNumericCellValue(), 0.0000000000000001); //A9: 1
		assertEquals(2, row7.getCell(3).getNumericCellValue(), 0.0000000000000001);	//D7: 2
		assertNull(sheet1.getRow(4)); //B5
		
		//C6: =A9
		Cell cellC6 = row6.getCell(2);
		CellValue valueC6 = _evaluator.evaluate(cellC6);
		assertEquals(1, valueC6.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC6.getCellType());
		testToFormulaString(cellC6, "A9");

		//A10: =C6
		Cell cellA10 = row10.getCell(0);
		CellValue valueA10 = _evaluator.evaluate(cellA10);
		assertEquals(1, valueA10.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA10.getCellType());
		testToFormulaString(cellA10, "C6");
		
		//A11: =B5
		Cell cellA11 = row11.getCell(0);
		CellValue valueA11 = _evaluator.evaluate(cellA11);
		assertEquals(0, valueA11.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA11.getCellType());
		testToFormulaString(cellA11, "B5");

		//A12: =D7
		Cell cellA12 = row12.getCell(0);
		CellValue valueA12 = _evaluator.evaluate(cellA12);
		assertEquals(2, valueA12.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA12.getCellType());
		testToFormulaString(cellA12, "D7");
		
		//merge B5:D7
		BookHelper.merge(sheet1, 4, 1, 6, 3, true);
		
		Row row5 = sheet1.getRow(4);
		assertEquals(Cell.CELL_TYPE_FORMULA, row5.getCell(1).getCellType()); //B5: null -> =A9 
		assertEquals(Cell.CELL_TYPE_BLANK, row5.getCell(2).getCellType()); //C5
		assertEquals(Cell.CELL_TYPE_BLANK, row5.getCell(3).getCellType()); //D5
		assertNull(row5.getCell(4)); //E5
		assertEquals(Cell.CELL_TYPE_BLANK, row6.getCell(1).getCellType()); //B6
		assertEquals(Cell.CELL_TYPE_BLANK, row6.getCell(2).getCellType()); //C6: =A9 -> blank
		assertEquals(Cell.CELL_TYPE_BLANK, row6.getCell(3).getCellType()); //D6
		assertNull(row6.getCell(4)); //E6
		assertEquals(Cell.CELL_TYPE_BLANK, row7.getCell(1).getCellType()); //B7
		assertEquals(Cell.CELL_TYPE_BLANK, row7.getCell(2).getCellType()); //C7
		assertEquals(Cell.CELL_TYPE_BLANK, row7.getCell(3).getCellType()); //D7: 2 -> blank
		assertNull(row7.getCell(4)); //E7
		
		//B5: =A9
		Cell cellB5 = row5.getCell(1);
		CellValue valueB5 = _evaluator.evaluate(cellB5);
		assertEquals(1, valueB5.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueB5.getCellType());
		testToFormulaString(cellB5, "A9");

		//A10: =C6 -> =B5
		_evaluator.notifySetFormula(cellA10);
		valueA10 = _evaluator.evaluate(cellA10);
		assertEquals(1, valueA10.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA10.getCellType());
		testToFormulaString(cellA10, "B5");
		
		//A12: =D7
		_evaluator.notifySetFormula(cellA12);
		valueA12 = _evaluator.evaluate(cellA12);
		assertEquals(0, valueA12.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA12.getCellType());
		testToFormulaString(cellA12, "D7");
		
		//A11: =B5 -> #REF!
		_evaluator.notifySetFormula(cellA11);
		valueA11 = _evaluator.evaluate(cellA11);
		assertEquals(Cell.CELL_TYPE_ERROR, valueA11.getCellType());
		assertEquals(ErrorConstants.ERROR_REF, valueA11.getErrorValue());

	}
	
	private void testToFormulaString(Cell cell, String expect) {
		HSSFEvaluationWorkbook evalbook = HSSFEvaluationWorkbook.create((HSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
