package zss.test;

import static org.junit.Assert.assertEquals;

import org.hamcrest.CoreMatchers;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ErrorCollector;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * For testing those user actions that can be performed on a cell, e.g. copy and paste.
 * 
 */
public class CellOperationTest extends SpreadsheetTestCaseBase{

	private DesktopAgent desktop;
	private ComponentAgent zss;
	private SpreadsheetAgent ssAgent;
	private Sheet sheet;
	
	@Rule
    public ErrorCollector collector = new ErrorCollector();
	
	public CellOperationTest(){

		desktop = Zats.newClient().connect("/operation.zul");
		zss = desktop.query("#operationTest");
		ssAgent = new SpreadsheetAgent(zss);
		sheet = zss.as(Spreadsheet.class).getSelectedSheet();
	}
	
	@Test
	public void testCopy() {

		//pre condition
		String srcText = Ranges.range(sheet, 2, 1).getCellEditText();
		String srcFormula = Ranges.range(sheet,3, 1).getCellEditText();
		
		ssAgent.copy(2, 7, 1, 2);
		ssAgent.paste(2, 4);
		
	
		String dstText = Ranges.range(sheet, 2, 4).getCellEditText();
		collector.checkThat(dstText, CoreMatchers.equalTo(srcText));
		String dstFormula = Ranges.range(sheet,3, 4).getCellEditText();
		collector.checkThat(dstFormula, CoreMatchers.equalTo(srcFormula));
		
	}
	
	/**
	 * set the background color of a cell below a merged cell,
	 *  the merged cell's bottom border will appear.
	 */
	@Test
	public void testChangeBackgroundIssue() {
		Range mergedCell = Ranges.range(sheet, 17, 1);
		assertEquals(CellStyle.BorderType.NONE, mergedCell.getCellStyle().getBorderBottom());
		//TODO how to check it's merged
		
		ssAgent.fillBackgroundColor(18, 1, "ff0000");
//		assertEquals(1, getCell(sheet,18,1).getCellStyle().getFillBackgroundColor());
		assertEquals(CellStyle.BorderType.NONE, mergedCell.getCellStyle().getBorderBottom());
	}
}
