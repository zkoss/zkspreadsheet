package zss.test.display;

import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.model.Worksheet;

import zss.test.SpreadsheetTestCaseBase;


/**
 * Test case base for the function "display Excel files".
 * @author Hawk
 *
 */
public class DisplayExcelTest extends SpreadsheetTestCaseBase{

	protected DesktopAgent desktop; 
	protected ComponentAgent zss ;
	protected Worksheet sheet;
	
	public DisplayExcelTest(String page){
		desktop = Zats.newClient().connect(page);
		zss = desktop.query("spreadsheet");
	}
	

}
