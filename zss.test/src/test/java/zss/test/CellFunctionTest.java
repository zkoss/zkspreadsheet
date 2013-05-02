package zss.test;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.zkoss.zats.mimic.ComponentAgent;
import org.zkoss.zats.mimic.DesktopAgent;
import org.zkoss.zats.mimic.Zats;
import org.zkoss.zats.mimic.operation.AuAgent;
import org.zkoss.zats.mimic.operation.AuData;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;

public class CellFunctionTest extends SpreadsheetTestCaseBase{

	@Test
	public void testCopy() {
		DesktopAgent desktop = Zats.newClient().connect("/display.zul");

		ComponentAgent zss = desktop.query("spreadsheet");
		Worksheet sheet = zss.as(Spreadsheet.class).getSelectedSheet();

		//pre condition
		String sourceText = Ranges.range(sheet, 0, 0).getEditText();
		
		//copy
		AuData copyEvent = new AuData("onZSSAction");
		copyEvent.setData("sheetId", "0").setData("tag", "toolbar").setData("act", "copy")
			.setData("tRow", 0).setData("lCol", 0)
			.setData("bRow", 0).setData("rCol", 0);
		zss.as(AuAgent.class).post(copyEvent);
		
		//paste
		AuData pasteEvent = new AuData("onZSSAction");
		pasteEvent.setData("sheetId", "0").setData("tag", "toolbar").setData("act", "paste")
			.setData("tRow", 0).setData("lCol", 1)
			.setData("bRow", 0).setData("rCol", 1);
		zss.as(AuAgent.class).post(pasteEvent);
		String destinationText = Ranges.range(sheet, 0, 1).getEditText();
		
		assertEquals(sourceText, destinationText);
	}
}
