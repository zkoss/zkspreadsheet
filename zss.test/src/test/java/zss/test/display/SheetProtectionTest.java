package zss.test.display;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetAgent;
import zss.test.SpreadsheetTestCaseBase;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheet "sheet-autofilter".
 * 
 * @author Hawk
 *
 */
public class SheetProtectionTest extends SpreadsheetTestCaseBase{

	private static DesktopAgent desktop; 
	private static ComponentAgent zss ;
	private static Worksheet sheet;
	
	@BeforeClass
	public static void initialize(){
		//sheet 0
		desktop = Zats.newClient().connect("/display.zul");
		
		zss = desktop.query("spreadsheet");
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("sheet-protection");
		sheet = zss.as(Spreadsheet.class).getBook().getWorksheet("sheet-protection");
		
	}
	

	@Test
	public void testProtection(){
		assertTrue(sheet.getProtect());
		assertTrue(getCell(sheet, 0, 0).getCellStyle().getLocked());
		//check editable cell
		assertFalse(getCell(sheet, 1, 0).getCellStyle().getLocked());
	}
	
}
