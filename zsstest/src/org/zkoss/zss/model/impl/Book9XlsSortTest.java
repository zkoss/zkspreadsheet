/**
 * 
 */
package org.zkoss.zss.model.impl;


import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

import java.io.InputStream;
import java.util.Set;

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
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.After;
import org.junit.Before;
import org.junit.Test;
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
public class Book9XlsSortTest {
	private Workbook _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book9.xls";
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
	public void testSortByColumns() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Row row2 = sheet1.getRow(1);
		Row row3 = sheet1.getRow(2);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		Row row8 = sheet1.getRow(7);
		Row row9 = sheet1.getRow(8);
		Row row10 = sheet1.getRow(9);
		Row row11 = sheet1.getRow(10);
		Row row12 = sheet1.getRow(11);
		Row row13 = sheet1.getRow(12);
		
		assertEquals(1, row1.getCell(0).getNumericCellValue(), 0.0000000000000001); //A1: 1
		assertEquals(2, row2.getCell(0).getNumericCellValue(), 0.0000000000000001);	//A2: 2
		assertEquals(3, row3.getCell(0).getNumericCellValue(), 0.0000000000000001); //A3: 3
		assertEquals("a", row4.getCell(0).getStringCellValue()); //A4: "a"
		assertEquals("b", row5.getCell(0).getStringCellValue()); //A5: "b"
		assertEquals("c", row6.getCell(0).getStringCellValue()); //A6: "c"
		assertEquals(ErrorConstants.ERROR_VALUE, row7.getCell(0).getErrorCellValue()); //A7: #VALUE!
		assertEquals(ErrorConstants.ERROR_REF, row8.getCell(0).getErrorCellValue()); //A8: #REF!
		assertNull(row9.getCell(0)); //A9: null
		assertEquals(ErrorConstants.ERROR_VALUE, row10.getCell(0).getErrorCellValue()); //A10: #VALUE!
		assertEquals(true, row11.getCell(0).getBooleanCellValue()); //A11: TRUE
		assertEquals(false, row12.getCell(0).getBooleanCellValue()); //A12: FALSE
		
		assertEquals(1, row1.getCell(1).getNumericCellValue(), 0.0000000000000001); //B1: 1
		assertEquals(2, row2.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B2: 2
		assertEquals(3, row3.getCell(1).getNumericCellValue(), 0.0000000000000001); //B3: 3
		assertEquals(4, row4.getCell(1).getNumericCellValue(), 0.0000000000000001); //B4: 4
		assertEquals(5, row5.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B5: 5
		assertEquals(6, row6.getCell(1).getNumericCellValue(), 0.0000000000000001); //B6: 6
		assertEquals(7, row7.getCell(1).getNumericCellValue(), 0.0000000000000001); //B7: 7
		assertEquals(8, row8.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B8: 8
		assertEquals(9, row9.getCell(1).getNumericCellValue(), 0.0000000000000001); //B9: 9
		assertEquals(10, row10.getCell(1).getNumericCellValue(), 0.0000000000000001); //B10: 10
		assertEquals(11, row11.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B11: 11
		assertEquals(12, row12.getCell(1).getNumericCellValue(), 0.0000000000000001); //B12: 12
		
		//C9: =B8
		Cell cellC9 = row9.getCell(2);
		CellValue valueC9 = _evaluator.evaluate(cellC9);
		assertEquals(8, valueC9.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC9.getCellType());
		testToFormulaString(cellC9, "B8");
		
		//C13: =B12
		Cell cellC13 = row13.getCell(2);
		CellValue valueC13 = _evaluator.evaluate(cellC13);
		assertEquals(12, valueC13.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC13.getCellType());
		testToFormulaString(cellC13, "B12");
		
		//Sort A1:C12
		Set<Ref>[] refs = BookHelper.sort(sheet1, 0, 0, 11, 2, Utils.getRange(sheet1, 0, 0).getRefs().iterator().next(), false, 
				null, 0, false, null, false, BookHelper.SORT_HEADER_NO, 0, false, false, 0, 
				BookHelper.SORT_NORMAL_DEFAULT, BookHelper.SORT_NORMAL_DEFAULT, BookHelper.SORT_NORMAL_DEFAULT);
		Set<Ref> last = refs[0];
		Set<Ref> all = refs[1];
		_evaluator.notifySetFormula(cellC13);

		assertEquals(1, row1.getCell(0).getNumericCellValue(), 0.0000000000000001); //A1: 1
		assertEquals(2, row2.getCell(0).getNumericCellValue(), 0.0000000000000001);	//A2: 2
		assertEquals(3, row3.getCell(0).getNumericCellValue(), 0.0000000000000001); //A3: 3
		assertEquals("a", row4.getCell(0).getStringCellValue()); //A4: "a"
		assertEquals("b", row5.getCell(0).getStringCellValue()); //A5: "b"
		assertEquals("c", row6.getCell(0).getStringCellValue()); //A6: "c"
		assertEquals(false, row7.getCell(0).getBooleanCellValue()); //A7: FALSE
		assertEquals(true, row8.getCell(0).getBooleanCellValue()); //A8: TRUE
		assertEquals(ErrorConstants.ERROR_VALUE, row9.getCell(0).getErrorCellValue()); //A9: #VALUE!
		assertEquals(ErrorConstants.ERROR_REF, row10.getCell(0).getErrorCellValue()); //A10: #REF!
		assertEquals(ErrorConstants.ERROR_VALUE, row11.getCell(0).getErrorCellValue()); //A11: #VALUE!
		assertNull(row12.getCell(0)); //A9: null
		
		assertEquals(1, row1.getCell(1).getNumericCellValue(), 0.0000000000000001); //B1: 1
		assertEquals(2, row2.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B2: 2
		assertEquals(3, row3.getCell(1).getNumericCellValue(), 0.0000000000000001); //B3: 3
		assertEquals(4, row4.getCell(1).getNumericCellValue(), 0.0000000000000001); //B4: 4
		assertEquals(5, row5.getCell(1).getNumericCellValue(), 0.0000000000000001);	//B5: 5
		assertEquals(6, row6.getCell(1).getNumericCellValue(), 0.0000000000000001); //B6: 6
		assertEquals(12, row7.getCell(1).getNumericCellValue(), 0.0000000000000001); //B7: 12
		assertEquals(11, row8.getCell(1).getNumericCellValue(), 0.0000000000000001); //B8: 11
		assertEquals(7, row9.getCell(1).getNumericCellValue(), 0.0000000000000001); //B9: 7
		assertEquals(8, row10.getCell(1).getNumericCellValue(), 0.0000000000000001); //B10: 8
		assertEquals(10, row11.getCell(1).getNumericCellValue(), 0.0000000000000001); //B11: 10
		assertEquals(9, row12.getCell(1).getNumericCellValue(), 0.0000000000000001); //B12: 9

		//C9 -> C12: =B8 -> B11
		Cell cellC12 = row12.getCell(2);
		CellValue valueC12 = _evaluator.evaluate(cellC12);
		assertEquals(10, valueC12.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC12.getCellType());
		testToFormulaString(cellC12, "B11");
		
		//C13: =B12
		cellC13 = row13.getCell(2);
		valueC13 = _evaluator.evaluate(cellC13);
		assertEquals(9, valueC13.getNumberValue(), 0.0000000000000001);
		assertEquals(Cell.CELL_TYPE_NUMERIC, valueC13.getCellType());
		testToFormulaString(cellC13, "B12");
	}
	
	@Test
	public void testSortByRows() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Row row15 = sheet1.getRow(14);
		Row row16 = sheet1.getRow(15);
		Row row17 = sheet1.getRow(16);
		
		assertEquals(1, row15.getCell(0).getNumericCellValue(), 0.0000000000000001); //A15: 1
		assertEquals(2, row15.getCell(1).getNumericCellValue(), 0.0000000000000001); //B15: 2
		assertEquals(3, row15.getCell(2).getNumericCellValue(), 0.0000000000000001); //C15: 3
		assertEquals("a", row15.getCell(3).getStringCellValue()); //D15: "a"
		assertEquals("b", row15.getCell(4).getStringCellValue()); //E15: "b"
		assertEquals("c", row15.getCell(5).getStringCellValue()); //F15: "c"
		assertEquals(ErrorConstants.ERROR_VALUE, row15.getCell(6).getErrorCellValue()); //G15: #VALUE!
		assertEquals(ErrorConstants.ERROR_REF, row15.getCell(7).getErrorCellValue()); //H15: #REF!
		assertNull(row15.getCell(8)); //I15: null
		assertEquals(ErrorConstants.ERROR_VALUE, row15.getCell(9).getErrorCellValue()); //J15: #VALUE!
		assertEquals(true, row15.getCell(10).getBooleanCellValue()); //K15: TRUE
		assertEquals(false, row15.getCell(11).getBooleanCellValue()); //L15: FALSE
		
		assertEquals("A", row16.getCell(0).getStringCellValue()); //A16: "A"
		assertEquals("B", row16.getCell(1).getStringCellValue()); //B16: "B"
		assertEquals("C", row16.getCell(2).getStringCellValue()); //C16: "C"
		assertEquals("D", row16.getCell(3).getStringCellValue()); //D16: "D"
		assertEquals("E", row16.getCell(4).getStringCellValue()); //E16: "E"
		assertEquals("F", row16.getCell(5).getStringCellValue()); //F16: "F"
		assertEquals("G", row16.getCell(6).getStringCellValue()); //G16: "G"
		assertEquals("H", row16.getCell(7).getStringCellValue()); //H16: "H"
		assertEquals("I", row16.getCell(8).getStringCellValue()); //I16: "I"
		assertEquals("J", row16.getCell(9).getStringCellValue()); //J16: "J"
		assertEquals("K", row16.getCell(10).getStringCellValue()); //K16: "K"
		assertEquals("L", row16.getCell(11).getStringCellValue()); //L16: "L"
		
		//I17: =H16
		Cell cellI17 = row17.getCell(8);
		CellValue valueI17 = _evaluator.evaluate(cellI17);
		assertEquals("H", valueI17.getStringValue());
		assertEquals(Cell.CELL_TYPE_STRING, valueI17.getCellType());
		testToFormulaString(cellI17, "H16");
		
		//M17: =L16
		Cell cellM17 = row17.getCell(12);
		CellValue valueM17 = _evaluator.evaluate(cellM17);
		assertEquals("L", valueM17.getStringValue());
		assertEquals(Cell.CELL_TYPE_STRING, valueM17.getCellType());
		testToFormulaString(cellM17, "L16");
		
		/*sort(Sheet sheet, int tRow, int lCol, int bRow, int rCol, 
				Ref key1, boolean desc1, Ref key2, int type, boolean desc2, Ref key3, boolean desc3, int header, int orderCustom,
				boolean matchCase, boolean sortByRows, int sortMethod, int dataOption1, int dataOption2, int dataOption3) */

		//Sort A15:L17
		Set<Ref>[] refs = BookHelper.sort(sheet1, 14, 0, 16, 11, Utils.getRange(sheet1, 14, 0).getRefs().iterator().next(), false, 
				null, 0, false, null, false, BookHelper.SORT_HEADER_NO, 0, false, true, 0, 
				BookHelper.SORT_NORMAL_DEFAULT, BookHelper.SORT_NORMAL_DEFAULT, BookHelper.SORT_NORMAL_DEFAULT);
		Set<Ref> last = refs[0];
		Set<Ref> all = refs[1];
		_evaluator.notifySetFormula(cellM17);

		assertEquals(1, row15.getCell(0).getNumericCellValue(), 0.0000000000000001); //A15: 1
		assertEquals(2, row15.getCell(1).getNumericCellValue(), 0.0000000000000001); //B15: 2
		assertEquals(3, row15.getCell(2).getNumericCellValue(), 0.0000000000000001); //C15: 3
		assertEquals("a", row15.getCell(3).getStringCellValue()); //D15: "a"
		assertEquals("b", row15.getCell(4).getStringCellValue()); //E15: "b"
		assertEquals("c", row15.getCell(5).getStringCellValue()); //F15: "c"
		assertEquals(false, row15.getCell(6).getBooleanCellValue()); //G15: FALSE
		assertEquals(true, row15.getCell(7).getBooleanCellValue()); //H15: TRUE
		assertEquals(ErrorConstants.ERROR_VALUE, row15.getCell(8).getErrorCellValue()); //I15: #VALUE!
		assertEquals(ErrorConstants.ERROR_REF, row15.getCell(9).getErrorCellValue()); //J15: #REF!
		assertEquals(ErrorConstants.ERROR_VALUE, row15.getCell(10).getErrorCellValue()); //K15: #VALUE!
		assertNull(row15.getCell(11)); //L15: null
		
		assertEquals("A", row16.getCell(0).getStringCellValue()); //A16: "A"
		assertEquals("B", row16.getCell(1).getStringCellValue()); //B16: "B"
		assertEquals("C", row16.getCell(2).getStringCellValue()); //C16: "C"
		assertEquals("D", row16.getCell(3).getStringCellValue()); //D16: "D"
		assertEquals("E", row16.getCell(4).getStringCellValue()); //E16: "E"
		assertEquals("F", row16.getCell(5).getStringCellValue()); //F16: "F"
		assertEquals("L", row16.getCell(6).getStringCellValue()); //G16: "L"
		assertEquals("K", row16.getCell(7).getStringCellValue()); //H16: "K"
		assertEquals("G", row16.getCell(8).getStringCellValue()); //I16: "G"
		assertEquals("H", row16.getCell(9).getStringCellValue()); //J16: "H"
		assertEquals("J", row16.getCell(10).getStringCellValue()); //K16: "J"
		assertEquals("I", row16.getCell(11).getStringCellValue()); //L16: "I"
		
		//I17 -> L17: =H16 -> K16
		Cell cellL17 = row17.getCell(11);
		CellValue valueL17 = _evaluator.evaluate(cellL17);
		assertEquals("J", valueL17.getStringValue());
		assertEquals(Cell.CELL_TYPE_STRING, valueL17.getCellType());
		testToFormulaString(cellL17, "K16");
		
		//M17: =L16
		cellM17 = row17.getCell(12);
		valueM17 = _evaluator.evaluate(cellM17);
		assertEquals("I", valueM17.getStringValue());
		assertEquals(Cell.CELL_TYPE_STRING, valueM17.getCellType());
		testToFormulaString(cellM17, "L16");
	}
	
	private void testToFormulaString(Cell cell, String expect) {
		HSSFEvaluationWorkbook evalbook = HSSFEvaluationWorkbook.create((HSSFWorkbook)_workbook);
		Ptg[] ptgs = BookHelper.getCellPtgs(cell);
		final String formula = FormulaRenderer.toFormulaString(evalbook, ptgs);
		assertEquals(expect, formula);
	}
}
