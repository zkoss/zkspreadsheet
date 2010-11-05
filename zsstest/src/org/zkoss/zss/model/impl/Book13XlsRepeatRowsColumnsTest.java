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
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.ExcelImporter;
import org.zkoss.zss.model.impl.HSSFBookImpl;
import org.zkoss.zss.ui.impl.Utils;

/**
 * Check if can correctly retrieve the repeat Rows and Columns
 * @author henrichen
 */
public class Book13XlsRepeatRowsColumnsTest {
	private Book _workbook;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "Book13.xls";
		final InputStream is = new ClassLocator().getResourceAsStream(filename);
		_workbook = new ExcelImporter().imports(is, filename);
		assertTrue(_workbook instanceof Book);
		assertTrue(_workbook instanceof HSSFBookImpl);
		assertTrue(_workbook instanceof HSSFWorkbook);
		assertEquals(filename, ((Book)_workbook).getBookName());
		assertEquals("Sheet1", _workbook.getSheetName(0));
		assertEquals("Sheet2", _workbook.getSheetName(1));
		assertEquals("Sheet3", _workbook.getSheetName(2));
		assertEquals("Sheet4", _workbook.getSheetName(3));
		assertEquals(0, _workbook.getSheetIndex("Sheet1"));
		assertEquals(1, _workbook.getSheetIndex("Sheet2"));
		assertEquals(2, _workbook.getSheetIndex("Sheet3"));
		assertEquals(3, _workbook.getSheetIndex("Sheet4"));
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_workbook = null;
	}
	
	@Test
	public void testgetRepeatingRowsAndColumns() {
		//$A:$C, $1:$4
		CellRangeAddress addr1 = _workbook.getRepeatingRowsAndColumns(0);
		assertEquals(0, addr1.getFirstColumn());
		assertEquals(2, addr1.getLastColumn());
		assertEquals(0, addr1.getFirstRow());
		assertEquals(3, addr1.getLastRow());
		
		//N/A, $6:$7
		CellRangeAddress addr2 = _workbook.getRepeatingRowsAndColumns(1);
		assertEquals(-1, addr2.getFirstColumn());
		assertEquals(-1, addr2.getLastColumn());
		assertEquals(5, addr2.getFirstRow());
		assertEquals(6, addr2.getLastRow());
		
		//$D:$E, $6:$7
		CellRangeAddress addr3 = _workbook.getRepeatingRowsAndColumns(2);
		assertEquals(3, addr3.getFirstColumn());
		assertEquals(4, addr3.getLastColumn());
		assertEquals(-1, addr3.getFirstRow());
		assertEquals(-1, addr3.getLastRow());
		
		//N/A,N/A
		CellRangeAddress addr4 = _workbook.getRepeatingRowsAndColumns(3);
		assertEquals(-1, addr4.getFirstColumn());
		assertEquals(-1, addr4.getLastColumn());
		assertEquals(-1, addr4.getFirstRow());
		assertEquals(-1, addr4.getLastRow());
	}
}
