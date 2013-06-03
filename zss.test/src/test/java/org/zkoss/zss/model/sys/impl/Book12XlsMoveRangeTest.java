/**
 * 
 */
package org.zkoss.zss.model.sys.impl;


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
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.ExcelImporter;
import org.zkoss.zss.model.sys.impl.HSSFBookImpl;

/**
 * Insert a row and check if the formula still work
 * @author henrichen
 */
public class Book12XlsMoveRangeTest {
	private XBook _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book12.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
		assertTrue(_workbook instanceof XBook);
		assertTrue(_workbook instanceof HSSFBookImpl);
		assertTrue(_workbook instanceof HSSFWorkbook);
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
	
	@Test
	public void testMoveE4F7_F4G7() { //right side move right 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move right from E4:F7 to F4:G7
		BookHelper.moveRange(sheet1, 3, 4, 6, 5, 0, 1);
		_evaluator.notifySetFormula(cellA1);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(6).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(6).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(6).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(6).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:G7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:G7)");
	}
	
	@Test
	public void testMoveC4D7_B4C7() { //left side move left 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from C4:D7 to B4:C7
		BookHelper.moveRange(sheet1, 3, 2, 6, 3, 0, -1);
		_evaluator.notifySetFormula(cellA1);
		
		assertEquals(1, row4.getCell(1).getNumericCellValue(), 0.0000000000000001); //B4: 1
		assertEquals(2, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(1).getNumericCellValue(), 0.0000000000000001); //B5: 5
		assertEquals(6, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(1).getNumericCellValue(), 0.0000000000000001); //B6: 9
		assertEquals(10, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(1).getNumericCellValue(), 0.0000000000000001); //B7: 13
		assertEquals(14, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7) -> =SUM(B4:F7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(B4:F7)");
	}
	
	@Test
	public void testMoveE4G7_D4F7() { //right side move left
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from E4:G7 to D4:F7
		BookHelper.moveRange(sheet1, 3, 4, 6, 6, 0, -1);
		_evaluator.notifySetFormula(cellA1);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(3, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 3
		assertEquals(4, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 4
		assertNull(row4.getCell(6)); //F4: null

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(7, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 7
		assertEquals(8, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 8
		assertNull(row5.getCell(6));  //F5: null
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(11, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 11
		assertEquals(12, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 12
		assertNull(row6.getCell(6)); //F6: null
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(15, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 15
		assertEquals(16, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 16
		assertNull(row7.getCell(6)); //F7: null
		
		//A1: =SUM(C4:G7) -> =SUM(C4:E7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(104, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:E7)");
	}
	
	@Test
	public void testMoveB4D7_C4E7() { //left side move right 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from B4:D7 to C4:E7
		BookHelper.moveRange(sheet1, 3, 1, 6, 3, 0, 1);
		_evaluator.notifySetFormula(cellA1);
		
		assertNull(row4.getCell(1)); //C4: null
		assertEquals(1, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 1
		assertEquals(2, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 2
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertNull(row5.getCell(1)); //C5: null
		assertEquals(5, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 5
		assertEquals(6, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 6
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertNull(row6.getCell(1)); //C6: null
		assertEquals(9, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(10, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertNull(row7.getCell(1)); //B7: 13
		assertEquals(13, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 13
		assertEquals(14, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 14
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7) -> =SUM(D4:F7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(100, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(D4:F7)");
	}
	
	@Test
	public void testMoveD4D7_F4F7() { //center side move right(within reference right edge) 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from D4:D7 to F4:F7
		BookHelper.moveRange(sheet1, 3, 3, 6, 3, 0, 2);
		_evaluator.notifySetFormula(cellA1);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertNull(row4.getCell(3));
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(2, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 2

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertNull(row5.getCell(3));
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(6, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 6
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertNull(row6.getCell(3));
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(10, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 10
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertNull(row7.getCell(3));
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(14, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 14
		
		
		//A1: =SUM(C4:F7) -> =SUM(C4:F7) //no change
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(96, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	}
	
	@Test
	public void testMoveF4G7_B4C7() { //right side move left and over original reference left column 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from F4:G7 to B4:C7
		BookHelper.moveRange(sheet1, 3, 5, 6, 6, 0, -4);
		_evaluator.notifySetFormula(cellA1);
		

		assertEquals(4, row4.getCell(1).getNumericCellValue(), 0.0000000000000001); //B4: 4
		assertNull(row4.getCell(2));
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3

		assertEquals(8, row5.getCell(1).getNumericCellValue(), 0.0000000000000001); //B5: 8
		assertNull(row5.getCell(2));
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		
		assertEquals(12, row6.getCell(1).getNumericCellValue(), 0.0000000000000001); //B6: 12
		assertNull(row6.getCell(2));
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		
		assertEquals(16, row7.getCell(1).getNumericCellValue(), 0.0000000000000001); //B7: 16
		assertNull(row7.getCell(2));
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		
		//A1: =SUM(C4:F7) -> =SUM(B4:E7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(108, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(B4:E7)");
	}
	
	@Test
	public void testMoveC6F8_C7F9() { //bottom side move down 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move right from C6:F8 to C7:F9
		BookHelper.moveRange(sheet1, 5, 2, 7, 5, 1, 0);
		_evaluator.notifySetFormula(cellA1);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 9
		assertEquals(10, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 10
		assertEquals(11, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 11
		assertEquals(12, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 12
		
		Row row8 = sheet1.getRow(7);
		assertEquals(13, row8.getCell(2).getNumericCellValue(), 0.0000000000000001); //C8: 13
		assertEquals(14, row8.getCell(3).getNumericCellValue(), 0.0000000000000001); //D8: 14
		assertEquals(15, row8.getCell(4).getNumericCellValue(), 0.0000000000000001); //E8: 15
		assertEquals(16, row8.getCell(5).getNumericCellValue(), 0.0000000000000001); //F8: 16
		
		//A1: =SUM(C4:F7) -> =SUM(C4:F8)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F8)");
	}
	
	@Test
	public void testMoveC3F5_C1F3() { //top side move up 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move up from C3:F4 to C1:F2
		BookHelper.moveRange(sheet1, 2, 2, 3, 5, -2, 0);
		_evaluator.notifySetFormula(cellA1);
		Row row2 = sheet1.getRow(1);
		assertEquals(1, row2.getCell(2).getNumericCellValue(), 0.0000000000000001); //C2: 1
		assertEquals(2, row2.getCell(3).getNumericCellValue(), 0.0000000000000001); //D2: 2
		assertEquals(3, row2.getCell(4).getNumericCellValue(), 0.0000000000000001); //E2: 3
		assertEquals(4, row2.getCell(5).getNumericCellValue(), 0.0000000000000001); //F2: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7) -> =SUM(C2:F7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C2:F7)");
	}
	
	@Test
	public void testMoveC6F8_C2F4() { //bottom side move top and over original reference top row  
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from C6:F8 to C2:F4
		BookHelper.moveRange(sheet1, 5, 2, 7, 5, -4, 0);
		_evaluator.notifySetFormula(cellA1);
		
		assertNull(row4.getCell(2));
		assertNull(row4.getCell(3));
		assertNull(row4.getCell(4));
		assertNull(row4.getCell(5));

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		Row row2 = sheet1.getRow(1);
		assertEquals(9, row2.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row2.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row2.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row2.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		Row row3 = sheet1.getRow(2);
		assertEquals(13, row3.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row3.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row3.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row3.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7) -> =SUM(C2:F5)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(126, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C2:F5)");
	}
	
	@Test
	public void testMoveC3F5_C4E7() { //top side move down 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from C3:F5 to C5:F7
		BookHelper.moveRange(sheet1, 2, 2, 4, 5, 2, 0);
		_evaluator.notifySetFormula(cellA1);
		
		assertEquals(1, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertNull(row4.getCell(2));
		assertNull(row4.getCell(3));
		assertNull(row4.getCell(4));
		assertNull(row4.getCell(5));
		
		assertNull(row5.getCell(2));
		assertNull(row5.getCell(3));
		assertNull(row5.getCell(4));
		assertNull(row5.getCell(5));
		
		//A1: =SUM(C4:F7) -> =SUM(C6:F7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(36, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C6:F7)");
	}
	
	@Test
	public void testMoveC6F6_C7F7() { //center side move down(within reference right edge) 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from C6:F6 to C7:F7
		BookHelper.moveRange(sheet1, 5, 2, 5, 5, 1, 0);
		_evaluator.notifySetFormula(cellA1);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertNull(row6.getCell(2));
		assertNull(row6.getCell(3));
		assertNull(row6.getCell(4));
		assertNull(row6.getCell(5));
		
		//A1: =SUM(C4:F7) -> =SUM(C4:F7) //no change
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(78, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	}
	
	@Test
	public void testMoveC7F9_C5F7() { //bottom side move up 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from C7:F9 to C5:F7
		BookHelper.moveRange(sheet1, 6, 2, 8, 5, -2, 0);
		_evaluator.notifySetFormula(cellA1);

		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(13, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 13
		assertEquals(14, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 14
		assertEquals(15, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 15
		assertEquals(16, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 16
		
		assertNull(row6.getCell(2));
		assertNull(row6.getCell(3));
		assertNull(row6.getCell(4));
		assertNull(row6.getCell(5));
		
		assertNull(row7.getCell(2));
		assertNull(row7.getCell(3));
		assertNull(row7.getCell(4));
		assertNull(row7.getCell(5));
		
		//A1: =SUM(C4:F7) -> =SUM(B4:E7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(68, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F5)");
	}
	
	@Test
	public void testMoveB2G8_C3H9() { //block move down-right 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from B2:G8 to C3:H9
		BookHelper.moveRange(sheet1, 1, 1, 7, 6, 1, 1);
		_evaluator.notifySetFormula(cellA1);

		assertNull(row5.getCell(2));
		assertEquals(1, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 1
		assertEquals(2, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 2
		assertEquals(3, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 3
		assertEquals(4, row5.getCell(6).getNumericCellValue(), 0.0000000000000001); //G4: 4

		assertNull(row6.getCell(2));
		assertEquals(5, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 5
		assertEquals(6, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 6
		assertEquals(7, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 7
		assertEquals(8, row6.getCell(6).getNumericCellValue(), 0.0000000000000001); //G4: 8
		
		assertNull(row7.getCell(2));
		assertEquals(9, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 9
		assertEquals(10, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 10
		assertEquals(11, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 11
		assertEquals(12, row7.getCell(6).getNumericCellValue(), 0.0000000000000001); //G4: 12
		
		Row row8 = sheet1.getRow(7);
		assertNull(row8.getCell(2));
		assertEquals(13, row8.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 13
		assertEquals(14, row8.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 14
		assertEquals(15, row8.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 15
		assertEquals(16, row8.getCell(6).getNumericCellValue(), 0.0000000000000001); //G4: 16
		
		assertNull(row4.getCell(2));
		assertNull(row4.getCell(3));
		assertNull(row4.getCell(4));
		assertNull(row4.getCell(5));
		
		//A1: =SUM(C4:F7) -> =SUM(B4:E7)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(D5:G8)");
	}
	
	@Test
	public void testMoveB2G8_B3F9() { //block move down-left 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from B2:G8 to C3:H9
		BookHelper.moveRange(sheet1, 1, 1, 7, 6, 1, -1);
		_evaluator.notifySetFormula(cellA1);

		assertNull(row5.getCell(5));
		assertEquals(1, row5.getCell(1).getNumericCellValue(), 0.0000000000000001); //D4: 1
		assertEquals(2, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //E4: 2
		assertEquals(3, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //F4: 3
		assertEquals(4, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //G4: 4

		assertNull(row6.getCell(5));
		assertEquals(5, row6.getCell(1).getNumericCellValue(), 0.0000000000000001); //D4: 5
		assertEquals(6, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //E4: 6
		assertEquals(7, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //F4: 7
		assertEquals(8, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //G4: 8
		
		assertNull(row7.getCell(5));
		assertEquals(9, row7.getCell(1).getNumericCellValue(), 0.0000000000000001); //D4: 9
		assertEquals(10, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //E4: 10
		assertEquals(11, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //F4: 11
		assertEquals(12, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //G4: 12
		
		Row row8 = sheet1.getRow(7);
		assertNull(row8.getCell(5));
		assertEquals(13, row8.getCell(1).getNumericCellValue(), 0.0000000000000001); //D4: 13
		assertEquals(14, row8.getCell(2).getNumericCellValue(), 0.0000000000000001); //E4: 14
		assertEquals(15, row8.getCell(3).getNumericCellValue(), 0.0000000000000001); //F4: 15
		assertEquals(16, row8.getCell(4).getNumericCellValue(), 0.0000000000000001); //G4: 16
		
		assertNull(row4.getCell(2));
		assertNull(row4.getCell(3));
		assertNull(row4.getCell(4));
		assertNull(row4.getCell(5));
		
		//A1: =SUM(C4:F7) -> =SUM(B5:E8)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(B5:E8)");
	}
	@Test
	public void testMoveB3G8_C2H7() { //block move up-right 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from B3:G8 to C2:H7
		BookHelper.moveRange(sheet1, 2, 1, 7, 6, -1, 1);
		_evaluator.notifySetFormula(cellA1);

		Row row3 = sheet1.getRow(2);
		assertNull(row3.getCell(2));
		assertEquals(1, row3.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 1
		assertEquals(2, row3.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 2
		assertEquals(3, row3.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 3
		assertEquals(4, row3.getCell(6).getNumericCellValue(), 0.0000000000000001); //G4: 4

		assertNull(row6.getCell(2));
		assertEquals(5, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 5
		assertEquals(6, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 6
		assertEquals(7, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 7
		assertEquals(8, row4.getCell(6).getNumericCellValue(), 0.0000000000000001); //G4: 8
		
		assertNull(row7.getCell(2));
		assertEquals(9, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 9
		assertEquals(10, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 10
		assertEquals(11, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 11
		assertEquals(12, row5.getCell(6).getNumericCellValue(), 0.0000000000000001); //G4: 12
		
		assertNull(row6.getCell(2));
		assertEquals(13, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 13
		assertEquals(14, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 14
		assertEquals(15, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 15
		assertEquals(16, row6.getCell(6).getNumericCellValue(), 0.0000000000000001); //G4: 16
		
		assertNull(row7.getCell(2));
		assertNull(row7.getCell(3));
		assertNull(row7.getCell(4));
		assertNull(row7.getCell(5));
		
		//A1: =SUM(C4:F7) -> =SUM(D3:G6)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(D3:G6)");
	}
	
	@Test
	public void testMoveB3G8_A2F7() { //block move up-left 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from B3:G8 to C2:H7
		BookHelper.moveRange(sheet1, 2, 1, 7, 6, -1, -1);
		_evaluator.notifySetFormula(cellA1);

		Row row3 = sheet1.getRow(2);
		assertNull(row3.getCell(5));
		assertEquals(1, row3.getCell(1).getNumericCellValue(), 0.0000000000000001); //D4: 1
		assertEquals(2, row3.getCell(2).getNumericCellValue(), 0.0000000000000001); //E4: 2
		assertEquals(3, row3.getCell(3).getNumericCellValue(), 0.0000000000000001); //F4: 3
		assertEquals(4, row3.getCell(4).getNumericCellValue(), 0.0000000000000001); //G4: 4

		assertNull(row6.getCell(5));
		assertEquals(5, row4.getCell(1).getNumericCellValue(), 0.0000000000000001); //D4: 5
		assertEquals(6, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //E4: 6
		assertEquals(7, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //F4: 7
		assertEquals(8, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //G4: 8
		
		assertNull(row7.getCell(5));
		assertEquals(9, row5.getCell(1).getNumericCellValue(), 0.0000000000000001); //D4: 9
		assertEquals(10, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //E4: 10
		assertEquals(11, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //F4: 11
		assertEquals(12, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //G4: 12
		
		assertNull(row6.getCell(5));
		assertEquals(13, row6.getCell(1).getNumericCellValue(), 0.0000000000000001); //D4: 13
		assertEquals(14, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //E4: 14
		assertEquals(15, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //F4: 15
		assertEquals(16, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //G4: 16
		
		assertNull(row7.getCell(2));
		assertNull(row7.getCell(3));
		assertNull(row7.getCell(4));
		assertNull(row7.getCell(5));
		
		//A1: =SUM(C4:F7) -> =SUM(B3:E6)
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(B3:E6)");
	}
	
	@Test
	public void testMoveD6G8_E7H9() { //partial block move down-right 
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		
		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertEquals(10, row6.getCell(3).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row6.getCell(4).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row6.getCell(5).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertEquals(14, row7.getCell(3).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7);
		Cell cellA1 = row1.getCell(0);
		CellValue valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(136, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	
		//move left from D6:G8 to E7:H9
		BookHelper.moveRange(sheet1, 5, 3, 7, 6, 1, 1);
		_evaluator.notifySetFormula(cellA1);

		assertEquals(1, row4.getCell(2).getNumericCellValue(), 0.0000000000000001); //C4: 1
		assertEquals(2, row4.getCell(3).getNumericCellValue(), 0.0000000000000001); //D4: 2
		assertEquals(3, row4.getCell(4).getNumericCellValue(), 0.0000000000000001); //E4: 3
		assertEquals(4, row4.getCell(5).getNumericCellValue(), 0.0000000000000001); //F4: 4

		assertEquals(5, row5.getCell(2).getNumericCellValue(), 0.0000000000000001); //C5: 5
		assertEquals(6, row5.getCell(3).getNumericCellValue(), 0.0000000000000001); //D5: 6
		assertEquals(7, row5.getCell(4).getNumericCellValue(), 0.0000000000000001); //E5: 7
		assertEquals(8, row5.getCell(5).getNumericCellValue(), 0.0000000000000001); //F5: 8
		
		assertEquals(9, row6.getCell(2).getNumericCellValue(), 0.0000000000000001); //C6: 9
		assertNull(row6.getCell(3));
		assertNull(row6.getCell(4));
		assertNull(row6.getCell(5));
		
		assertEquals(13, row7.getCell(2).getNumericCellValue(), 0.0000000000000001); //C7: 13
		assertNull(row7.getCell(3));
		
		assertEquals(10, row7.getCell(4).getNumericCellValue(), 0.0000000000000001); //D6: 10
		assertEquals(11, row7.getCell(5).getNumericCellValue(), 0.0000000000000001); //E6: 11
		assertEquals(12, row7.getCell(6).getNumericCellValue(), 0.0000000000000001); //F6: 12
		
		Row row8 = sheet1.getRow(7);
		assertNull(row8.getCell(2));
		assertNull(row8.getCell(3));
		assertEquals(14, row8.getCell(4).getNumericCellValue(), 0.0000000000000001); //D7: 14
		assertEquals(15, row8.getCell(5).getNumericCellValue(), 0.0000000000000001); //E7: 15
		assertEquals(16, row8.getCell(6).getNumericCellValue(), 0.0000000000000001); //F7: 16
		
		//A1: =SUM(C4:F7) -> =SUM(C4:F7) //no change
		cellA1 = row1.getCell(0);
		valueA1 = _evaluator.evaluate(cellA1);
		assertEquals(79, valueA1.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueA1.getCellType());
		testToFormulaString(cellA1, "SUM(C4:F7)");
	}
	
	private void testToFormulaString(Cell cell, String expect) {
		HSSFEvaluationWorkbook evalbook = HSSFEvaluationWorkbook.create((HSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
