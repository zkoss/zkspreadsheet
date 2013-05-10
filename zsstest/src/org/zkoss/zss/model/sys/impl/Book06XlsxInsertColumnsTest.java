/**
 * 
 */
package org.zkoss.zss.model.sys.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.ss.formula.ptg.Ptg;
import org.zkoss.poi.hssf.usermodel.HSSFEvaluationWorkbook;
import org.zkoss.poi.hssf.usermodel.HSSFWorkbook;
import org.zkoss.poi.hssf.util.HSSFColor;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.ss.usermodel.FormulaEvaluator;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.xssf.usermodel.XSSFColor;
import org.zkoss.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.ExcelImporter;
import org.zkoss.zss.model.sys.impl.XSSFBookImpl;

/**
 * Insert a column and check if the formula still work
 * @author henrichen
 */
public class Book06XlsxInsertColumnsTest {
	private XBook _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book6.xlsx";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
		assertTrue(_workbook instanceof XBook);
		assertTrue(_workbook instanceof XSSFBookImpl);
		assertTrue(_workbook instanceof XSSFWorkbook);
		assertEquals(filename, ((XBook)_workbook).getBookName());
		assertEquals("Sheet1", _workbook.getSheetName(0));
		assertEquals("Sheet2", _workbook.getSheetName(1));
		assertEquals("Sheet3", _workbook.getSheetName(2));
		assertEquals(0, _workbook.getSheetIndex("Sheet1"));
		assertEquals(1, _workbook.getSheetIndex("Sheet2"));
		assertEquals(2, _workbook.getSheetIndex("Sheet3"));

		_evaluator = ((XBook)_workbook).getFormulaEvaluator();
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_workbook = null;
		_evaluator = null;
	}

	private void assertColors(byte[] expect, byte[] real) {
		assertEquals(expect.length, real.length);
		for (int j = 0; j < expect.length; ++j) {
			assertEquals(expect[j], real[j]);
		}
	}
	@Test
	public void testInsertColumnC() {
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row2 = sheet1.getRow(1);
		Row row3 = sheet1.getRow(2);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		assertEquals(1, row1.getCell(5).getNumericCellValue(), 0.0000000000000001); //F1: 1
		assertEquals(2, row2.getCell(5).getNumericCellValue(), 0.0000000000000001);	//F2: 2
		assertEquals(3, row3.getCell(5).getNumericCellValue(), 0.0000000000000001); //F3: 3
		assertEquals(7, row4.getCell(0).getNumericCellValue(), 0.0000000000000001); //A4: 7
		assertEquals(9, row1.getCell(2).getNumericCellValue(), 0.0000000000000001); //C1: 9
		assertEquals(11, row1.getCell(3).getNumericCellValue(), 0.0000000000000001); //D1: 11
		
		byte[] REDColor = new byte[] {(byte)0xff, 0, 0};
		byte[] YELLOWColor = new byte[] {(byte)0xff, (byte)0xff, 0};
		byte[] ffg = ((XSSFColor)row1.getCell(1).getCellStyle().getFillForegroundColorColor()).getRgb();
		assertColors(REDColor, ffg);
		byte[] fbg = ((XSSFColor)row1.getCell(1).getCellStyle().getFillBackgroundColorColor()).getRgb();
		assertColors(YELLOWColor, fbg);
		
		//A1: =SUM(F1:F3)
		Cell cellA1 = row1.getCell(0); //A1: =SUM(F1:F3)
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(6, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(F1:F3)");
		
		//A2: =SUM(D2:XFD2) XFD: 16384
		Cell cellA2 = row2.getCell(0);
		CellValue valueA2 = _evaluator.evaluate(cellA2);
		assertEquals(2, valueA2.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA2.getCellType());
		testToFormulaString(cellA2, "SUM(D2:XFD2)");

		//A3: =SUM(XFC3:XFD3) XFC: 16383, XFD: 16384
		Cell cellA3 = row3.getCell(0); 
		CellValue valueA3 = _evaluator.evaluate(cellA3);
		assertEquals(0, valueA3.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA3.getCellType());
		testToFormulaString(cellA3, "SUM(XFC3:XFD3)");
		
		//A5: =SUM(C1:D1)
		Cell cellA5 = row5.getCell(0);
		CellValue valueA5 = _evaluator.evaluate(cellA5);
		assertEquals(20, valueA5.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA5.getCellType());
		testToFormulaString(cellA5, "SUM(C1:D1)");
		
		//A6: =SUM(B1:D1)
		Cell cellA6 = row6.getCell(0);
		CellValue valueA6 = _evaluator.evaluate(cellA6);
		assertEquals(20, valueA6.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA6.getCellType());
		testToFormulaString(cellA6, "SUM(B1:D1)");
		
		//H1: =SUM(F1:F3)
		Cell cellH1 = row1.getCell(7); 
		CellValue valueH1 = _evaluator.evaluate(cellH1);
		assertEquals(6, valueH1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueH1.getCellType());
		testToFormulaString(cellH1, "SUM(F1:F3)");

		//H4: =A4
		Cell cellH4 = row4.getCell(7); 
		CellValue valueH4 = _evaluator.evaluate(cellH4);
		assertEquals(7, valueH4.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueH4.getCellType());
		testToFormulaString(cellH4, "A4");
		
		//Insert before column C
		BookHelper.insertColumns(sheet1, 2, 1, XRange.FORMAT_LEFTABOVE);
		_evaluator.notifySetFormula(cellA1);
		_evaluator.notifySetFormula(cellA2);
		_evaluator.notifySetFormula(cellA3);
		_evaluator.notifySetFormula(cellA5);
		_evaluator.notifySetFormula(cellA6);
		
		assertEquals(sheet1.getColumnWidth(1), sheet1.getColumnWidth(2)); //column c width == column b width
		assertColors(REDColor, ((XSSFColor)row1.getCell(1).getCellStyle().getFillForegroundColorColor()).getRgb());
		assertColors(YELLOWColor, ((XSSFColor)row1.getCell(1).getCellStyle().getFillBackgroundColorColor()).getRgb());
		assertColors(REDColor, ((XSSFColor)row1.getCell(2).getCellStyle().getFillForegroundColorColor()).getRgb());
		assertColors(YELLOWColor, ((XSSFColor)row1.getCell(2).getCellStyle().getFillBackgroundColorColor()).getRgb());
		
		assertEquals(1, row1.getCell(6).getNumericCellValue(), 0.0000000000000001); //G1: 1
		assertEquals(2, row2.getCell(6).getNumericCellValue(), 0.0000000000000001);	//G2: 2
		assertEquals(3, row3.getCell(6).getNumericCellValue(), 0.0000000000000001); //G3: 3
		assertEquals(7, row4.getCell(0).getNumericCellValue(), 0.0000000000000001); //A4: 7

		//A1: =SUM(G1:G3)
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(6, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(G1:G3)");
		
		//A2: =SUM(E2:XFD2) XFD: 16384
		valueA2 = _evaluator.evaluate(cellA2);
		assertEquals(2, valueA2.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA2.getCellType());
		testToFormulaString(cellA2, "SUM(E2:XFD2)");

		//A3: =SUM(XFD3:XFD3) XFD: 16384
		cellA3 = row3.getCell(0); //A3: =SUM(IU3:IV3) 
		valueA6 = _evaluator.evaluate(cellA3);
		assertEquals(0, valueA6.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA6.getCellType());
		testToFormulaString(cellA3, "SUM(XFD3:XFD3)");
		
		//A5: =SUM(C1:D1) -> =SUM(D1:E1)
		cellA5 = row5.getCell(0);
		valueA5 = _evaluator.evaluate(cellA5);
		assertEquals(20, valueA5.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA5.getCellType());
		testToFormulaString(cellA5, "SUM(D1:E1)");
		
		//A6: =SUM(B1:D1) -> =SUM(B1:E1)
		cellA6 = row6.getCell(0);
		valueA6 = _evaluator.evaluate(cellA6);
		assertEquals(20, valueA6.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA6.getCellType());
		testToFormulaString(cellA6, "SUM(B1:E1)");
		
		//I1: =SUM(G1:G3)
		Cell cellI1 = row1.getCell(8); 
		CellValue valueI1 = _evaluator.evaluate(cellI1);
		assertEquals(6, valueI1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueI1.getCellType());
		testToFormulaString(cellI1, "SUM(G1:G3)");

		//I4: =A4
		Cell cellI4 = row4.getCell(8); 
		CellValue valueI4 = _evaluator.evaluate(cellI4);
		assertEquals(7, valueI4.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueI4.getCellType());
		testToFormulaString(cellI4, "A4");
		
	}
	
	@Test
	public void testInsertRangeC1_C4() {
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row2 = sheet1.getRow(1);
		Row row3 = sheet1.getRow(2);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		assertEquals(1, row1.getCell(5).getNumericCellValue(), 0.0000000000000001); //F1: 1
		assertEquals(2, row2.getCell(5).getNumericCellValue(), 0.0000000000000001);	//F2: 2
		assertEquals(3, row3.getCell(5).getNumericCellValue(), 0.0000000000000001); //F3: 3
		assertEquals(7, row4.getCell(0).getNumericCellValue(), 0.0000000000000001); //A4: 7
		assertEquals(9, row1.getCell(2).getNumericCellValue(), 0.0000000000000001); //C1: 9
		assertEquals(11, row1.getCell(3).getNumericCellValue(), 0.0000000000000001); //D1: 11
		
		//A1: =SUM(F1:F3)
		Cell cellA1 = row1.getCell(0); //A1: =SUM(F1:F3)
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(6, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(F1:F3)");
		
		//A2: =SUM(D2:XFD2) XFD: 16384
		Cell cellA2 = row2.getCell(0);
		CellValue valueA2 = _evaluator.evaluate(cellA2);
		assertEquals(2, valueA2.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA2.getCellType());
		testToFormulaString(cellA2, "SUM(D2:XFD2)");

		//A3: =SUM(XFC3:XFD3) XFC: 16383, XFD: 16384
		Cell cellA3 = row3.getCell(0); //A3: =SUM(XFC3:XFD3) 
		CellValue valueA3 = _evaluator.evaluate(cellA3);
		assertEquals(0, valueA3.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA3.getCellType());
		testToFormulaString(cellA3, "SUM(XFC3:XFD3)");
		
		//A5: =SUM(C1:D1)
		Cell cellA5 = row5.getCell(0);
		CellValue valueA5 = _evaluator.evaluate(cellA5);
		assertEquals(20, valueA5.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA5.getCellType());
		testToFormulaString(cellA5, "SUM(C1:D1)");
		
		//A6: =SUM(B1:D1)
		Cell cellA6 = row6.getCell(0);
		CellValue valueA6 = _evaluator.evaluate(cellA6);
		assertEquals(20, valueA6.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA6.getCellType());
		testToFormulaString(cellA6, "SUM(B1:D1)");
		
		//H1: =SUM(F1:F3)
		Cell cellH1 = row1.getCell(7); 
		CellValue valueH1 = _evaluator.evaluate(cellH1);
		assertEquals(6, valueH1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueH1.getCellType());
		testToFormulaString(cellH1, "SUM(F1:F3)");

		//H4: =A4
		Cell cellH4 = row4.getCell(7); 
		CellValue valueH4 = _evaluator.evaluate(cellH4);
		assertEquals(7, valueH4.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueH4.getCellType());
		testToFormulaString(cellH4, "A4");
		
		//Insert C1:C4
		BookHelper.insertRange(sheet1, 0, 2, 3, 2, true, XRange.FORMAT_LEFTABOVE);
		_evaluator.notifySetFormula(cellA1);
		_evaluator.notifySetFormula(cellA2);
		_evaluator.notifySetFormula(cellA3);
		_evaluator.notifySetFormula(cellA5);
		_evaluator.notifySetFormula(cellA6);
		
		assertEquals(1, row1.getCell(6).getNumericCellValue(), 0.0000000000000001); //G1: 1
		assertEquals(2, row2.getCell(6).getNumericCellValue(), 0.0000000000000001);	//G2: 2
		assertEquals(3, row3.getCell(6).getNumericCellValue(), 0.0000000000000001); //G3: 3
		assertEquals(7, row4.getCell(0).getNumericCellValue(), 0.0000000000000001); //A4: 7
		assertEquals(9, row1.getCell(3).getNumericCellValue(), 0.0000000000000001); //D1: 9
		assertEquals(11, row1.getCell(4).getNumericCellValue(), 0.0000000000000001); //E1: 11

		//A1: =SUM(G1:G3)
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(6, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(G1:G3)");
		
		//A2: =SUM(E2:XFD2) XFD: 16384
		valueA2 = _evaluator.evaluate(cellA2);
		assertEquals(2, valueA2.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA2.getCellType());
		testToFormulaString(cellA2, "SUM(E2:XFD2)");

		//A3: =SUM(XFC3:XFD3) -> =SUM(XFD3:XFD3) XFD: 16384
		cellA3 = row3.getCell(0); 
		valueA6 = _evaluator.evaluate(cellA3);
		assertEquals(0, valueA6.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA6.getCellType());
		testToFormulaString(cellA3, "SUM(XFD3:XFD3)");
		
		//A5: =SUM(C1:D1) -> =SUM(D1:E1)
		cellA5 = row5.getCell(0);
		valueA5 = _evaluator.evaluate(cellA5);
		assertEquals(20, valueA5.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA5.getCellType());
		testToFormulaString(cellA5, "SUM(D1:E1)");
		
		//A6: =SUM(B1:D1) -> =SUM(B1:E1)
		cellA6 = row6.getCell(0);
		valueA6 = _evaluator.evaluate(cellA6);
		assertEquals(20, valueA6.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA6.getCellType());
		testToFormulaString(cellA6, "SUM(B1:E1)");
		
		//I1: =SUM(G1:G3)
		Cell cellI1 = row1.getCell(8); 
		CellValue valueI1 = _evaluator.evaluate(cellI1);
		assertEquals(6, valueI1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueI1.getCellType());
		testToFormulaString(cellI1, "SUM(G1:G3)");

		//I4: =A4
		Cell cellI4 = row4.getCell(8); 
		CellValue valueI4 = _evaluator.evaluate(cellI4);
		assertEquals(7, valueI4.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueI4.getCellType());
		testToFormulaString(cellI4, "A4");
		
	}
	
	private void testToFormulaString(Cell cell, String expect) {
		XSSFEvaluationWorkbook evalbook = XSSFEvaluationWorkbook.create((XSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
