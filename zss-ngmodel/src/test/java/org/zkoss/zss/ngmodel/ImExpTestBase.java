package org.zkoss.zss.ngmodel;

import static org.junit.Assert.assertEquals;

import java.text.DateFormat;
import java.util.Locale;

import org.zkoss.zss.ngmodel.NCellStyle.BorderType;

/**
 * importer & exporter test case base.
 * migrate the common test case here.
 * @author kuro, hawk
 *
 */
public class ImExpTestBase {
	
	/**
	 * Test sheet in book
	 * @param book
	 */
	protected void sheetTest(NBook book) {
		assertEquals(7, book.getNumOfSheet());

		NSheet sheet1 = book.getSheet(0);
		assertEquals("Value", sheet1.getSheetName());
		assertEquals(21, sheet1.getDefaultRowHeight());
		assertEquals(72, sheet1.getDefaultColumnWidth());
		
		NSheet sheet2 = book.getSheet(1);
		assertEquals("Style", sheet2.getSheetName());
		NSheet sheet3 = book.getSheet(2);
		assertEquals("NamedRange", sheet3.getSheetName());
	}
	
	/**
	 * Test cell value
	 */
	protected void cellValueTest(NBook book) {
		NSheet sheet = book.getSheet(0);
		//text
		NRow row = sheet.getRow(0);
		assertEquals(NCell.CellType.STRING, sheet.getCell(0,1).getType());
		assertEquals("B1", sheet.getCell(0,1).getStringValue());
		assertEquals("C1", sheet.getCell(0,2).getStringValue());
		assertEquals("D1", sheet.getCell(0,3).getStringValue());
		
		//number
		NRow row1 = sheet.getRow(1);
		assertEquals(NCell.CellType.NUMBER, sheet.getCell(1,1).getType());
		assertEquals(123, sheet.getCell(1,1).getNumberValue().intValue());
		assertEquals(123.45, sheet.getCell(1,2).getNumberValue().doubleValue(), 0.01);
		
		//date
		NRow row2 = sheet.getRow(2);
		assertEquals(NCell.CellType.NUMBER, sheet.getCell(2,1).getType());
		assertEquals(41618, sheet.getCell(2,1).getNumberValue().intValue());
		assertEquals("Dec 10, 2013", DateFormat.getDateInstance(DateFormat.MEDIUM, Locale.US).format(sheet.getCell(2,1).getDateValue()));
		assertEquals(0.61, sheet.getCell(2,2).getNumberValue().doubleValue(), 0.01);
		assertEquals("2:44:10 PM", DateFormat.getTimeInstance (DateFormat.MEDIUM, Locale.US).format(sheet.getCell(2,2).getDateValue()));
		
		//formula
		NRow row3 = sheet.getRow(3);
		assertEquals(NCell.CellType.FORMULA, sheet.getCell(3,1).getType());
		assertEquals("SUM(10,20)", sheet.getCell(3,1).getFormulaValue());
		assertEquals("ISBLANK(B1)", sheet.getCell(3,2).getFormulaValue());
		assertEquals("B1", sheet.getCell(3,3).getFormulaValue());
		
		//error
		NRow row4 = sheet.getRow(4);
		assertEquals(NCell.CellType.ERROR, sheet.getCell(4,1).getType());
		assertEquals(ErrorValue.INVALID_NAME, sheet.getCell(4,1).getErrorValue().getCode());
		assertEquals(ErrorValue.INVALID_VALUE, sheet.getCell(4,2).getErrorValue().getCode());
		
		//blank
		NRow row5 = sheet.getRow(5);
		assertEquals(NCell.CellType.BLANK, sheet.getCell(5,1).getType());
		assertEquals("", sheet.getCell(5,1).getStringValue());
	}
	
	/**
	 * FIXME undone
	 * @param book
	 */
	protected void borderTypeTest(NBook book) {


	}
}
