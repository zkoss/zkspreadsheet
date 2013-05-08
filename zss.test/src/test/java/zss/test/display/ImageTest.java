package zss.test.display;

import static org.junit.Assert.assertEquals;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;

import zss.test.SpreadsheetAgent;
import zss.test.SpreadsheetTestCaseBase;


/**
 * Test case for the function "display Excel files".
 * Testing for the sheets "chart-image".
 * 
 * @author Hawk
 *
 */
public class ImageTest extends SpreadsheetTestCaseBase{

	private static DesktopAgent desktop; 
	private static ComponentAgent zss ;
	private static Spreadsheet spreadsheet;
	private static Worksheet sheet;
	
	
	@BeforeClass
	public static void initialize(){
		desktop = Zats.newClient().connect("/display.zul");
		zss = desktop.query("spreadsheet");
		SpreadsheetAgent ssAgent = new SpreadsheetAgent(zss);
		//select the sheet first or the chart won't be initialized
		ssAgent.selectSheet("chart-image");
		spreadsheet = zss.as(Spreadsheet.class);
		sheet = spreadsheet.getBook().getWorksheet("chart-image");
		
	}

	@Test
	public void testPicture(){
		assertEquals(1, sheet.getPictures().size());
		assertEquals(0, sheet.getPictures().get(0).getPreferredSize().getCol1());
		assertEquals(1, sheet.getPictures().get(0).getPreferredSize().getCol2());
		assertEquals(1, sheet.getPictures().get(0).getPreferredSize().getRow1());
		assertEquals(4, sheet.getPictures().get(0).getPreferredSize().getRow2());
		
	}
	
	@Test
	public void testCharts(){
		assertEquals(12, sheet.getCharts().size());
	}
}
