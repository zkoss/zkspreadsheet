package org.zkoss.zss;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Sheet;

public class SpreadsheetTestCaseBase {
	@BeforeClass
	public static void init() {
		Zats.init("./src/main/webapp"); 
	}

	@AfterClass
	public static void end() {
		Zats.end();
	}

	@After
	public void after() {
		Zats.cleanup();
	}
	
	protected CellStyle getCellStyle(Sheet sheet, int row, int col) {
		return Ranges.range(sheet, row, col).getCellStyle();
	}

	/*
	protected XSSFColor getCellFontColor(Spreadsheet spreadsheet, int row, int column) {
		//FIXME do not use BookHelper
		return (XSSFColor)BookHelper.getFontColor(spreadsheet.getBook(),
			spreadsheet.getBook().getFontAt(Utils.getCell(spreadsheet.getSelectedSheet(), row, column).getCellStyle().getFontIndex()));
	}
	*/

}
