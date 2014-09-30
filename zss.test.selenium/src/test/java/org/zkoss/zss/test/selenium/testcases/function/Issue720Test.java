package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue720Test extends ZSSTestCase {
	@Ignore("A bit complicated")
	@Test
	public void testZSS725() throws Exception {
		
	}
	
	@Ignore("Need export")
	@Test
	public void testZSS728() throws Exception {
		
	}
	
	@Test
	public void testZSS729() throws Exception {
		getTo("/issue3/729-validation.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		click(jq("@button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		CellWidget A1 = sheetCtrl.getCell(0, 0);
		click(A1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(ZSStyle.DROPDOWN_BTN));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("abc", jq(".zsdv-popup-cave").text());
		click(jq("@button:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		CellWidget B1 = sheetCtrl.getCell(0, 1);
		JQuery dropdown = jq(ZSStyle.DROPDOWN_BTN);
		click(B1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("def", jq(".zsdv-popup-cave").text());
		click(jq("@button:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(A1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse(dropdown.exists());
		click(B1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse(dropdown.exists());
	}
}
