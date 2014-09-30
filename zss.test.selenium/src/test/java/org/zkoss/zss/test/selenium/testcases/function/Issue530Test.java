package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue530Test extends ZSSTestCase {
	
	@Ignore("Don't know how to assert")
	@Test
	public void testZSS533() throws Exception {
	}
	
	@Test
	public void testZSS534() throws Exception {
		getTo("/issue3/534-fontsize.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		
		CellWidget cellB1 = spreadsheet.getSheetCtrl().getCell("B1");
		
		spreadsheet.focus();
		click(cellB1);
		waitForTime(Setup.getTimeoutL0());
		assertEquals(cellB1.getText(), jq(".zsfontsize input").val());
		click(jq(ZSStyle.SHEET_TAB + ":eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq(ZSStyle.SHEET_TAB + ":eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals(cellB1.getText(), jq(".zsfontsize input").val());
	}
}
