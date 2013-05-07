package zss.test;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.impl.Utils;

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
	
	//encapsulate access methods for future API change
	protected Cell getCell(Worksheet sheet, int row, int col) {
		return Utils.getCell(sheet, row, col);
	}

	/*
	protected XSSFColor getCellFontColor(Spreadsheet spreadsheet, int row, int column) {
		//FIXME do not use BookHelper
		return (XSSFColor)BookHelper.getFontColor(spreadsheet.getBook(),
			spreadsheet.getBook().getFontAt(Utils.getCell(spreadsheet.getSelectedSheet(), row, column).getCellStyle().getFontIndex()));
	}
	*/

}
