/**
 * 
 */
package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.ss.usermodel.FormulaEvaluator;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;

/**
 * Insert a row and check if the formula still work
 * @author henrichen
 */
public class Book05XlsxDeleteRowsTest {
	private Book _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book5.xlsx";
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
	public void testDelete4_6() {
		Worksheet sheet1 = _workbook.getWorksheet("Sheet1");
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
		
		//6 merge area
		assertEquals(6, sheet1.getNumMergedRegions());
		for (int j = 0; j < 6; ++j) {
			CellRangeAddress rng = sheet1.getMergedRegion(j);
			switch(j) {
			case 0:
				assertEquals("E6:F8", rng.formatAsString());
				break;
			case 1:
				assertEquals("E10:F12", rng.formatAsString());
				break;
			case 2:
				assertEquals("G5:H5", rng.formatAsString());
				break;
			case 3:
				assertEquals("E3:F4", rng.formatAsString());
				break;
			case 4:
				assertEquals("I3:J7", rng.formatAsString());
				break;
			case 5:
				assertEquals("E14:F15", rng.formatAsString());
				break;
			}
		}
	
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
		assertNull(row10);
		
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
		
		//5 merge area
		assertEquals(5, sheet1.getNumMergedRegions()); //G5:H5 unmerged
		for (int j = 0; j < 5; ++j) {
			CellRangeAddress rng = sheet1.getMergedRegion(j);
			switch(j) {
			case 0:
				assertEquals("E4:F5", rng.formatAsString()); //E6:F8 -> E4:F5
				break;
			case 1:
				assertEquals("E7:F9", rng.formatAsString()); //E10:F12 -> E7:F9
				break;
			case 2:
				assertEquals("E3:F3", rng.formatAsString()); //E3:F4 -> E3:F3
				break;
			case 3:
				assertEquals("I3:J4", rng.formatAsString()); //I3:J7 -> I3:J4
				break;
			case 4:
				assertEquals("E11:F12", rng.formatAsString()); //E14:F15 -> E11:F12
			}
		}
	}
	
	@Test
	public void testDeleteA4_J6() {
		Worksheet sheet1 = _workbook.getWorksheet("Sheet1");
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
		
		//6 merge area
		assertEquals(6, sheet1.getNumMergedRegions());
		for (int j = 0; j < 6; ++j) {
			CellRangeAddress rng = sheet1.getMergedRegion(j);
			switch(j) {
			case 0:
				assertEquals("E6:F8", rng.formatAsString());
				break;
			case 1:
				assertEquals("E10:F12", rng.formatAsString());
				break;
			case 2:
				assertEquals("G5:H5", rng.formatAsString());
				break;
			case 3:
				assertEquals("E3:F4", rng.formatAsString());
				break;
			case 4:
				assertEquals("I3:J7", rng.formatAsString());
				break;
			case 5:
				assertEquals("E14:F15", rng.formatAsString());
				break;
			}
		}
	
		//remove A4:J6
		BookHelper.deleteRange(sheet1, 3, 0, 5, 9, false);
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
		
		//5 merge area
		assertEquals(5, sheet1.getNumMergedRegions()); //G5:H5 unmerged
		for (int j = 0; j < 5; ++j) {
			CellRangeAddress rng = sheet1.getMergedRegion(j);
			switch(j) {
			case 0:
				assertEquals("E4:F5", rng.formatAsString()); //E6:F8 -> E4:F5
				break;
			case 1:
				assertEquals("E7:F9", rng.formatAsString()); //E10:F12 -> E7:F9
				break;
			case 2:
				assertEquals("E3:F3", rng.formatAsString()); //E3:F4 -> E3:F3
				break;
			case 3:
				assertEquals("I3:J4", rng.formatAsString()); //I3:J7 -> I3:J4
				break;
			case 4:
				assertEquals("E11:F12", rng.formatAsString()); //E14:F15 -> E11:F12
			}
		}
	}
	
	private void testToFormulaString(Cell cell, String expect) {
		XSSFEvaluationWorkbook evalbook = XSSFEvaluationWorkbook.create((XSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
