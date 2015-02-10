package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue730Test extends ZSSTestCase{
	
	@Test
	public void testZSS736() throws Exception {
		final String text = " test";
		getTo("issue3/736-sheet-copy.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		click(ctrl.getCell("A2"));
		sendKeys(jq("body"), text);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		rightClick(jq(".zssheettab:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-copySheet"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		// sheet 4
		assertTrue(jq(".zssheettab:eq(3)").text().trim().equals("Sheet1 (2)"));
		sheetFunction().gotoTab(4);
		waitUntilProcessEnd(Setup.getTimeoutL1());
		assertTrue(jq(".zsrow:eq(1) .zscell:eq(0)").text().trim().equals(text.trim()));
		
		rightClick(jq(".zssheettab:eq(3)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-copySheet"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		// sheet 5
		assertTrue(jq(".zssheettab:eq(4)").text().trim().equals("Sheet1 (3)"));
		sheetFunction().gotoTab(5);
		waitUntilProcessEnd(Setup.getTimeoutL1());
		assertTrue(jq(".zsrow:eq(1) .zscell:eq(0)").text().equals(text.trim()));
		
		rightClick(jq(".zssheettab:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-copySheet"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
			
		// sheet 6
		assertTrue(jq(".zssheettab:eq(5)").text().trim().equals("Sheet1 (4)"));
		sheetFunction().gotoTab(6);
		waitUntilProcessEnd(Setup.getTimeoutL1());
		assertTrue(jq(".zsrow:eq(1) .zscell:eq(0)").text().equals(text.trim()));
			
		rightClick(jq(".zssheettab:eq(3)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-renameSheet"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(jq(".zssheettab-rename-textbox:visible"), "Sheet1 (15)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(ctrl.getCell("A2"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		rightClick(jq(".zssheettab:eq(3)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-copySheet"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		// sheet 7
		assertTrue(jq(".zssheettab:eq(6)").text().trim().equals("Sheet1 (16)"));
		sheetFunction().gotoTab(7);
		waitUntilProcessEnd(Setup.getTimeoutL1());
		assertTrue(jq(".zsrow:eq(1) .zscell:eq(0)").text().equals(text.trim()));
		
	}
	
	@Ignore("IME issue")
	@Test
	public void testZSS737() throws Exception {
		
	}
}
