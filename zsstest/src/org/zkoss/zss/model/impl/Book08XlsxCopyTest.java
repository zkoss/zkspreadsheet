/**
 * 
 */
package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.apache.poi.hssf.record.constant.ErrorConstant;
import org.apache.poi.hssf.record.formula.Ptg;
import org.apache.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.ErrorConstants;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;

/**
 * Insert a row and check if the formula still work
 * @author henrichen
 */
public class Book08XlsxCopyTest {
	private Workbook _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book8.xlsx";
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
	public void testCopyCell() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row2 = sheet1.getRow(1);
		Row row3 = sheet1.getRow(2);
		assertEquals(1, row1.getCell(0).getNumericCellValue(), 0.0000000000000001); //A1: 1
		assertEquals(2, row1.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B1: 2
		assertEquals(3, row2.getCell(0).getNumericCellValue(), 0.0000000000000001); //A2: 3
		assertEquals(4, row2.getCell(1).getNumericCellValue(), 0.0000000000000001); //B2: 4
		
		//C3: =A1+7
		Cell cellC3 = row3.getCell(2);
		CellValue valueC3 = _evaluator.evaluate(cellC3);
		assertEquals(8, valueC3.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC3.getCellType());
		testToFormulaString(cellC3, "A1+7");
		
		//Copy cell (C3 -> D4)
		BookHelper.copyCell(cellC3, sheet1, 3, 3, Range.PASTE_ALL, Range.PASTEOP_NONE);
		_evaluator.notifySetFormula(cellC3);

		//A1,A2,A2,B2 stay as is
		assertEquals(1, row1.getCell(0).getNumericCellValue(), 0.0000000000000001); //A1: 1
		assertEquals(2, row1.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B1: 2
		assertEquals(3, row2.getCell(0).getNumericCellValue(), 0.0000000000000001); //A2: 3
		assertEquals(4, row2.getCell(1).getNumericCellValue(), 0.0000000000000001); //B2: 4
		
		//C3 stay as is
		valueC3 = _evaluator.evaluate(cellC3);
		assertEquals(8, valueC3.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC3.getCellType());
		testToFormulaString(cellC3, "A1+7");
		
		//D4: =B2+7
		Row row4 = sheet1.getRow(3);
		Cell cellD4 = row4.getCell(3);
		CellValue valueD4 = _evaluator.evaluate(cellD4);
		assertEquals(11, valueD4.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueD4.getCellType());
		testToFormulaString(cellD4, "B2+7");
	}
	
	@Test
	public void testCopyCellRefError2() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row2 = sheet1.getRow(1);
		Row row3 = sheet1.getRow(2);
		assertEquals(1, row1.getCell(0).getNumericCellValue(), 0.0000000000000001); //A1: 1
		assertEquals(2, row1.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B1: 2
		assertEquals(3, row2.getCell(0).getNumericCellValue(), 0.0000000000000001); //A2: 3
		assertEquals(4, row2.getCell(1).getNumericCellValue(), 0.0000000000000001); //B2: 4
		
		//C3: =A1+7
		Cell cellC3 = row3.getCell(2);
		CellValue valueC3 = _evaluator.evaluate(cellC3);
		assertEquals(8, valueC3.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC3.getCellType());
		testToFormulaString(cellC3, "A1+7");
		
		//Copy cell (C3 -> C2)
		BookHelper.copyCell(cellC3, sheet1, 1, 2, Range.PASTE_ALL, Range.PASTEOP_NONE);
		_evaluator.notifySetFormula(cellC3);

		//A1,A2,A2,B2 stay as is
		assertEquals(1, row1.getCell(0).getNumericCellValue(), 0.0000000000000001); //A1: 1
		assertEquals(2, row1.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B1: 2
		assertEquals(3, row2.getCell(0).getNumericCellValue(), 0.0000000000000001); //A2: 3
		assertEquals(4, row2.getCell(1).getNumericCellValue(), 0.0000000000000001); //B2: 4
		
		//C3 stay as is
		valueC3 = _evaluator.evaluate(cellC3);
		assertEquals(8, valueC3.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC3.getCellType());
		testToFormulaString(cellC3, "A1+7");
		
		//C2: #REF!
		Cell cellC2 = row2.getCell(2);
		CellValue valueC2 = _evaluator.evaluate(cellC2);
		assertEquals(ErrorConstants.ERROR_REF, valueC2.getErrorValue());
		assertEquals(Cell.CELL_TYPE_ERROR, valueC2.getCellType());
		testToFormulaString(cellC2, "#REF!+7");
	}
	
	@Test
	public void testCopyCellRefError3() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row2 = sheet1.getRow(1);
		Row row3 = sheet1.getRow(2);
		assertEquals(1, row1.getCell(0).getNumericCellValue(), 0.0000000000000001); //A1: 1
		assertEquals(2, row1.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B1: 2
		assertEquals(3, row2.getCell(0).getNumericCellValue(), 0.0000000000000001); //A2: 3
		assertEquals(4, row2.getCell(1).getNumericCellValue(), 0.0000000000000001); //B2: 4
		assertEquals(5, row1.getCell(4).getNumericCellValue(), 0.0000000000000001); //E1: 5
		assertEquals(6, row1.getCell(5).getNumericCellValue(), 0.0000000000000001); //F1: 6
		
		//D3: =SUM(E1:F1)
		Cell cellD3 = row3.getCell(3);
		CellValue valueD3 = _evaluator.evaluate(cellD3);
		assertEquals(11, valueD3.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueD3.getCellType());
		testToFormulaString(cellD3, "SUM(E1:F1)");
		
		//Copy cell (D3 -> XFD3)
		BookHelper.copyCell(cellD3, sheet1, 2, 16383, Range.PASTE_ALL, Range.PASTEOP_NONE);
		_evaluator.notifySetFormula(cellD3);

		//A1,A2,A2,B2,E1,F1 stay as is
		assertEquals(1, row1.getCell(0).getNumericCellValue(), 0.0000000000000001); //A1: 1
		assertEquals(2, row1.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B1: 2
		assertEquals(3, row2.getCell(0).getNumericCellValue(), 0.0000000000000001); //A2: 3
		assertEquals(4, row2.getCell(1).getNumericCellValue(), 0.0000000000000001); //B2: 4
		assertEquals(5, row1.getCell(4).getNumericCellValue(), 0.0000000000000001); //E1: 5
		assertEquals(6, row1.getCell(5).getNumericCellValue(), 0.0000000000000001); //F1: 6
		
		//D3 stay as is
		valueD3 = _evaluator.evaluate(cellD3);
		assertEquals(11, valueD3.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueD3.getCellType());
		testToFormulaString(cellD3, "SUM(E1:F1)");
		
		//XFD3: #REF!
		Cell cellXFD3 = row3.getCell(16383);
		CellValue valueIV3 = _evaluator.evaluate(cellXFD3);
		assertEquals(ErrorConstants.ERROR_REF, valueIV3.getErrorValue());
		assertEquals(Cell.CELL_TYPE_ERROR, valueIV3.getCellType());
		testToFormulaString(cellXFD3, "SUM(#REF!)");
	}
	
	private void testToFormulaString(Cell cell, String expect) {
		XSSFEvaluationWorkbook evalbook = XSSFEvaluationWorkbook.create((XSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
