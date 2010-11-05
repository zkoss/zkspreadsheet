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
import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.engine.Ref;
import org.zkoss.zss.engine.impl.CellRefImpl;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.FormatText;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;
import org.zkoss.zss.ui.impl.Utils;

/**
 * Check if the Text format correctly.
 * @author henrichen
 */
public class Book14XlsTextFormatTest {
	private Workbook _workbook;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book14.xls";
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
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	@Test
	public void testCellNumberFormatter() {
		Sheet sheet1 = _workbook.getSheet("Sheet1");
		Row row2 = sheet1.getRow(1);
		Row row3 = sheet1.getRow(2);
		Row row4 = sheet1.getRow(3);
		Row row5 = sheet1.getRow(4);
		Row row6 = sheet1.getRow(5);
		Row row7 = sheet1.getRow(6);
		Row row8 = sheet1.getRow(7);
		Row row9 = sheet1.getRow(8);
		Row row10 = sheet1.getRow(9);
	
		testOneRow(row2); // 1/4
		testOneRow(row3); // 21/25
		testOneRow(row4); // 312/943
		testOneRow(row5); // 1/2
		testOneRow(row6); // 2/4
		testOneRow(row7); // 4/8
		testOneRow(row8); // 8/16
		testOneRow(row9); // 3/10
		testOneRow(row10); // 30/100
		
		testZero(row2, 2, "0    "); //0 + 4 spaces
		testZero(row3, 2, "0      "); //0 + 6 spaces
		testZero(row4, 2, "0        "); //0 + 8 spaces
		testZero(row5, 2, "0    "); //0 + 4 spaces
		testZero(row6, 2, "0    "); //0 + 4 spaces
		testZero(row7, 2, "0    "); //0 + 4 spaces
		testZero(row8, 2, "0      "); //0 + 6 spaces
		testZero(row9, 2, "0     "); //0 + 5 spaces
		testZero(row10, 2, "0       "); //0 + 7 spaces

		testZero(row5, 3, "0/2"); // 0/2
		testZero(row6, 3, "0/4"); // 0/4
		testZero(row7, 3, "0/8"); // 0/8
		testZero(row8, 3, " 0/16"); // 1 spaces + 0/16
		testZero(row9, 3, "0/10"); // 0/10
		testZero(row10, 3, " 0/100"); // 1 spaces + 0/100
		testZero(row2, 3, "0/1"); // 0/1
		testZero(row3, 3, " 0/1 "); // 1 spaces + 0/1 + 1 spaces
		testZero(row4, 3, "  0/1  "); // 2 spaces + 0/1 + 2 spaces
	}

	private void testOneRow(Row row) {
		Cell cellA = row.getCell(0);
		FormatText a = BookHelper.getFormatText(cellA);
		assertTrue("Should be a rich text string", a.isRichTextString());
		assertTrue("Should not be a Format cell", !a.isCellFormatResult());
		
		Cell cellB = row.getCell(1);
		FormatText b = BookHelper.getFormatText(cellB);
		assertTrue(!b.isRichTextString());
		assertTrue(b.isCellFormatResult());
		
		assertEquals(a.getRichTextString().getString(), b.getCellFormatResult().text);
	}
	
	private void testZero(Row row, int colindex, String expect) {
		Cell cell = row.getCell(colindex);
		FormatText c = BookHelper.getFormatText(cell);
		assertTrue(!c.isRichTextString());
		assertTrue(c.isCellFormatResult());
		
		assertEquals(expect, c.getCellFormatResult().text);
	}

}
