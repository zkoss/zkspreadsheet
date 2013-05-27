package zss.test;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 */
public class SpreadsheetAgentTest extends SpreadsheetTestCaseBase{

	private DesktopAgent desktop;
	private SpreadsheetAgent ssAgent;
	private ComponentAgent spreadsheet;
	
	public SpreadsheetAgentTest(){

		desktop = Zats.newClient().connect("/display.zul");
		spreadsheet = desktop.query("spreadsheet");
		ssAgent = new SpreadsheetAgent(spreadsheet);
	}
	
	@Test
	public void testSelectSheet() {
		ssAgent.selectSheet("cell-border");
		Sheet selectedSheet = spreadsheet.as(Spreadsheet.class).getSelectedSheet();
		assertEquals("cell-border", selectedSheet.getSheetName());
		
		ssAgent.selectSheet("row");
		selectedSheet = spreadsheet.as(Spreadsheet.class).getSelectedSheet();
		assertEquals("row", selectedSheet.getSheetName());
	}

	
}
