package zss.test.display;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.usermodel.AutoFilter;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
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
public class SheetTest extends SpreadsheetTestCaseBase{

	private static DesktopAgent desktop; 
	private static ComponentAgent zss ;
//	private static Spreadsheet spreadsheet ;
	private static Worksheet sheet;
	
	@BeforeClass
	public static void initialize(){
		//sheet 0
		desktop = Zats.newClient().connect("/display.zul");
		
		zss = desktop.query("spreadsheet");
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		ssAgent.selectSheet("sheet-autofilter");
		sheet = zss.as(Spreadsheet.class).getBook().getWorksheet("sheet-autofilter");
		
	}
	

	@Test
	public void testAutofilter(){
		AutoFilter autoFilter = sheet.getAutoFilter(); 
		assertNotNull(autoFilter);
		assertTrue(autoFilter.getFilterColumn(1).isOn());
		assertEquals(1, autoFilter.getFilterColumn(1).getFilters().size());
		assertEquals("Sunday", autoFilter.getFilterColumn(1).getFilters().get(0));
		for (int r1=1 ; r1<=6 ; r1++){
			assertTrue(isHiddenRow(sheet, r1));
		}
		for (int r2=8 ; r2<=13 ; r2++){
			assertTrue(isHiddenRow(sheet, r2));
		}
	}
	
}
