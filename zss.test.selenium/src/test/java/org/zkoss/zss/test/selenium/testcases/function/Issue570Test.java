package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue570Test extends ZSSTestCase {
	
	@Test
	public void testZSS576() throws Exception {
		getTo("issue3/576-sheetProtect.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		// 1. Focus and selection should be allowed only inside the greed area in Sheet "Test1".
		dragAndDrop(ctrl.getCell("A4").$n(), ctrl.getCell("C6").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 3"), false);
		
		dragAndDrop(ctrl.getCell("E5").$n(), ctrl.getCell("F6").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 4"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.bottom === 5"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.left === 4"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.right === 5"), true);
		
		// 2. Make a selection and click "Clear All" or "Clear Style", the selection should be unselectable.
		click(jq(".zstbtn-clear:visible .zstbtn-cave"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-clearAll:visible"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(ctrl.getCell("G7"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		dragAndDrop(ctrl.getCell("E5").$n(), ctrl.getCell("F6").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 4"), false);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.bottom === 5"), false);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.left === 4"), false);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.right === 5"), false);
		
		// 3. Switch to "Test2" then go back to "Test1", and do step 1.
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetFunction().gotoTab(1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		dragAndDrop(ctrl.getCell("A4").$n(), ctrl.getCell("C6").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 3"), false);
		
		click(ctrl.getCell("E7"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 6"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.bottom === 6"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.left === 4"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.right === 4"), true);
		
		// 4. Sort and Insert Hyperlink should only work in the greed area in Sheet "Test2".
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		dragAndDrop(ctrl.getCell("B6").$n(), ctrl.getCell("B10").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zstbtn-sortAndFilter .zstbtn-cave"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-sortAscending"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-window .z-button"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(ctrl.getCell("B6").getText(), "B");
		
		dragAndDrop(ctrl.getCell("F7").$n(), ctrl.getCell("F11").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zstbtn-sortAndFilter .zstbtn-cave"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-sortAscending"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(ctrl.getCell("F7").getText(), "A");
		
		// 5. Move focus should not crash or slower in Sheet "All Unlock cell".
		sheetFunction().gotoTab(4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		dragAndDrop(ctrl.getCell("E5").$n(), ctrl.getCell("F6").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 4"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.bottom === 5"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.left === 4"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.right === 5"), true);
		
		dragAndDrop(ctrl.getCell("A1").$n(), ctrl.getCell("F6").$n());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.top === 0"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.bottom === 5"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.left === 0"), true);
		assertEquals(
				eval("return jq('@spreadsheet').zk.$().sheetCtrl.selArea.lastRange.right === 5"), true);
	}
	
	@Test
	public void testZSS577() throws Exception {
		getTo("issue3/577-chart.zul");
		assertEquals(jq(".zswidget-chart").length(), 2);
		assertEquals(sheetChartUtil().getValue(".zswidget-chart:eq(0)", 0, 0, "y"), "1");
		assertEquals(sheetChartUtil().getValue(".zswidget-chart:eq(0)", 0, 1, "y"), "3");
		
		assertEquals(sheetChartUtil().getValue(".zswidget-chart:eq(1)", 0, 0, "y"), "1");
		assertEquals(sheetChartUtil().getValue(".zswidget-chart:eq(1)", 0, 1, "y"), "3");
	}
	
	@Test
	public void testZSS579() throws Exception {
		getTo("/issue3/579-undo-toolbarbutton.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellC4 = sheetCtrl.getCell("C4");
		EditorWidget editor = sheetCtrl.getInlineEditor();
		WebElement fontBold = jq(ZSStyle.FONT_BOLD_BTN).toWebElement();
		spreadsheet.focus();
		click(cellC4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(ZSStyle.FONT_BOLD_BTN));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(fontBold.getAttribute("class").contains(ZSStyle.ACTIVE_BTN.getName()));
		editor.toWebElement().sendKeys(Keys.chord(Keys.CONTROL, "z"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse(fontBold.getAttribute("class").contains(ZSStyle.ACTIVE_BTN.getName()));
	}
}
