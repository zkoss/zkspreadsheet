package zss.test.display;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetTestCaseBase;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "cell-reference".
 * 
 * @author Hawk
 *
 */
public class CellReferenceTest extends SpreadsheetTestCaseBase{


	@Test
	public void testCrossSheetReference(){
		DesktopAgent desktop = Zats.newClient().connect("/display.zul");
		
		ComponentAgent zss = desktop.query("spreadsheet");
		Sheet sheet = zss.as(Spreadsheet.class).getBook().getSheet("cell-reference");
		
//		assertEquals("The first row is freezed.", Ranges.range(sheet, 3, 2).getText().getString());
		assertEquals("The first row is freezed.", Ranges.range(sheet, 3, 2).getCellValueHelper().getFormatText());
		assertEquals("9", Ranges.range(sheet, 4, 2).getCellValueHelper().getFormatText());
	}
	
	/*
	 * FIXME the application page is not ready, ignore this case.
	 */
	@Test
	public void testExternalFileReference(){
		DesktopAgent desktop = Zats.newClient().connect("/external-reference.zul");
		
		ComponentAgent zss = desktop.query("#dst");
		Sheet sheet = zss.as(Spreadsheet.class).getBook().getSheet("cell-reference");
		
		assertEquals("ZK", Ranges.range(sheet, 0, 2).getCellValueHelper().getFormatText());
		assertEquals("10", Ranges.range(sheet, 1, 2).getCellValueHelper().getFormatText());
	}
}
