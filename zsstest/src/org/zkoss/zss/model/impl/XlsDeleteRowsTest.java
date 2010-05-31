/**
 * 
 */
package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;

/**
 * Insert a row and check if the formula still work
 * @author henrichen
 */
public class XlsDeleteRowsTest {
	private Workbook _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book5.xls";
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
	public void testDeleteOneRow() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		Row row8 = sheet1.getRow(7);
		Row row9 = sheet1.getRow(8);
		Row row10 = sheet1.getRow(9);
		assertEquals(1, row5.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(2, row5.getCell(2).getNumericCellValue(), 0.0000000000000001);
		assertEquals(3, row6.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(5, row7.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(7, row8.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(9, row9.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(11, row10.getCell(1).getNumericCellValue(), 0.0000000000000001);

		//A1: =SUM(B5:C5)
		Row row1 = sheet1.getRow(0);
		Cell cellA1 = row1.getCell(0); //A1: =SUM(B5:C5)
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(3, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(B5:C5)");

		//B1: =SUM(B5:B10)
		Cell cellB1 = row1.getCell(1); //B1: =SUM(B5:B10)
		CellValue valueB1 = _evaluator.evaluate(cellB1);
		assertEquals(36, valueB1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueB1.getCellType());
		testToFormulaString(cellB1, "SUM(B5:B10)");
		
		//C1: =SUM(B8:B10)
		Cell cellC1 = row1.getCell(2); //C1: =SUM(B8:B10)
		CellValue valueC1 = _evaluator.evaluate(cellC1);
		assertEquals(27, valueC1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC1.getCellType());
		testToFormulaString(cellC1, "SUM(B8:B10)");
	
		//remove rows 4 ~ 6
		BookHelper.deleteRows(sheet1, 3, 3); //remove rows 4 ~ 6
		_evaluator.notifySetFormula(cellA1);
		_evaluator.notifySetFormula(cellB1);
		_evaluator.notifySetFormula(cellC1);
		
		//B4: 5, B5:7, C5: empty, B6: 9, B7: 11, row 8 ~ row 10 empty
		Row row4 = sheet1.getRow(3);
		row5 = sheet1.getRow(4);
		row6 = sheet1.getRow(5);
		row7 = sheet1.getRow(6);
		
		row8 = sheet1.getRow(7);
		row9 = sheet1.getRow(8);
		row10 = sheet1.getRow(9);
		
		assertNull(row5.getCell(2)); //C5 not exist
		assertNull(row8.getCell(1));
		assertNull(row9.getCell(1));
		assertNull(row10.getCell(1));
		
		assertEquals(5, row4.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(7, row5.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(9, row6.getCell(1).getNumericCellValue(), 0.0000000000000001);
		assertEquals(11, row7.getCell(1).getNumericCellValue(), 0.0000000000000001);
		
		//C1: =SUM(B5:B7)
		valueC1 = _evaluator.evaluate(cellC1);
		assertEquals(27, valueC1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC1.getCellType());
		testToFormulaString(cellC1, "SUM(B5:B7)");
		
		//B1: =SUM(B4:B7)
		valueB1 = _evaluator.evaluate(cellB1);
		assertEquals(32, valueB1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueB1.getCellType());
		testToFormulaString(cellB1, "SUM(B4:B7)");
		
		//A1: =SUM(#REF!)
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(Cell.CELL_TYPE_ERROR, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(#REF!)");
	}
	
	private void testToFormulaString(Cell cell, String expect) {
		HSSFEvaluationWorkbook evalbook = HSSFEvaluationWorkbook.create((HSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
