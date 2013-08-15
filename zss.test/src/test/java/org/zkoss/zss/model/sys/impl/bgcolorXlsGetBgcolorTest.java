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
import org.zkoss.poi.hssf.util.HSSFColor;
import org.zkoss.poi.ss.formula.FormulaRenderer;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.CellValue;
import org.zkoss.poi.ss.usermodel.Color;
import org.zkoss.poi.ss.usermodel.FormulaEvaluator;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.xssf.usermodel.XSSFColor;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.util.resource.ClassLocator;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.ExcelImporter;
import org.zkoss.zss.model.sys.impl.HSSFBookImpl;
import org.zkoss.zss.model.sys.impl.XSSFBookImpl;
import org.zkoss.zss.ui.impl.Utils;

/**
 * Read and check if cell background color is correct.
 * @author henrichen
 */
public class bgcolorXlsGetBgcolorTest {
	private XBook _workbook;
	private FormulaEvaluator _evaluator;

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		final String filename = "bgcolor.xlsx";
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
	
	@Test
	public void testGetBgColor() {
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		assertEquals("FFFF0000", ((XSSFColor)row1.getCell(0).getCellStyle().getFillForegroundColorColor()).getARGBHex()); //A1: Red
	}
	
	@Test
	public void testSetBgColor() {
		XSheet sheet1 = _workbook.getWorksheet("Sheet1");
		Row row1 = sheet1.getRow(0);
		Cell A1 = BookHelper.getOrCreateCell(sheet1, 0, 0);
		CellStyle newCellStyle = _workbook.createCellStyle();
		newCellStyle.cloneStyleFrom(A1.getCellStyle());
		final Color bsColor = BookHelper.HTMLToColor(_workbook, "#00FF00");
		BookHelper.setFillForegroundColor(newCellStyle, bsColor);
		BookHelper.setCellStyle(sheet1, 0, 0, 0, 0, newCellStyle);
		assertEquals("FF00FF00", ((XSSFColor)row1.getCell(0).getCellStyle().getFillForegroundColorColor()).getARGBHex()); //A1: Green
	}
}
