package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue480Test extends ZSSTestCase {
	@Test
	public void testZSS485() throws Exception {
		getTo("/issue3/485-insert-row-copy.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		SheetCtrlWidget sheetCtrl = getSpreadsheet().getSheetCtrl();
		assertEquals("1", sheetCtrl.getCell("A4").getText());
	    assertEquals("2", sheetCtrl.getCell("B4").getText());
	    assertEquals("3", sheetCtrl.getCell("C4").getText());
	    assertEquals("4", sheetCtrl.getCell("A5").getText());
	    assertEquals("5", sheetCtrl.getCell("B5").getText());
	    assertEquals("6", sheetCtrl.getCell("C5").getText());
	    assertEquals("7", sheetCtrl.getCell("A6").getText());
	    assertEquals("8", sheetCtrl.getCell("B6").getText());
	    assertEquals("9", sheetCtrl.getCell("C6").getText());
	    assertEquals("a", sheetCtrl.getCell("A7").getText());
	    assertEquals("b", sheetCtrl.getCell("B7").getText());
	    assertEquals("c", sheetCtrl.getCell("C7").getText());
	    assertEquals("d", sheetCtrl.getCell("A8").getText());
	    assertEquals("e", sheetCtrl.getCell("B8").getText());
	    assertEquals("f", sheetCtrl.getCell("C8").getText());
	    assertEquals("g", sheetCtrl.getCell("A9").getText());
	    assertEquals("h", sheetCtrl.getCell("B9").getText());
	    assertEquals("i", sheetCtrl.getCell("C9").getText());
	}
	
	@Test
	public void testZSS488() throws Exception {
		getTo("issue3/488-delete-first-column.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		click(".z-button:eq(2)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(ctrl.getCell("A1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 0"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.bottom === 0"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.left === 0"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.right === 0"), true);
		
		assertEquals(ctrl.getFormulabarEditor().getValue(), "B1");
		
		getTo("issue3/488-delete-first-column.zul");
		ss = focusSheet();
		ctrl = ss.getSheetCtrl();
		
		click(".z-button:eq(3)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(".z-button:eq(4)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(".z-button:eq(5)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		// directly click on A1 will got B1 clicked :<
		click(ctrl.getCell("A2"));
		click(ctrl.getCell("A1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 0"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.bottom === 0"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.left === 0"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.right === 0"), true);
		
		assertEquals(ctrl.getFormulabarEditor().getValue(), "A2");
	}
}
