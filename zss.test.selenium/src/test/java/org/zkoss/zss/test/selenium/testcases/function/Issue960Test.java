package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import java.util.Map;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue960Test extends ZSSTestCase {
	
	@Test
	public void testZSS962() throws Exception {
		getTo("/issue3/962-subtotal-ignore-hidden.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
		for(int col = 6; col <= 7; col++) {
			for(int row = 11; row <= 18; row++) {
				assertEquals("TRUE", ctrl.getCell(row, col).getText());
			}
		}
	}
	
	@Test
	public void testZSS963() throws Exception {
		getTo("/issue3/963-text-indent-toolbar.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
		click(ctrl.getCell("A1"));
		waitForTime(Setup.getTimeoutL0());
		sendKeys(editor, "test indent");
		waitForTime(Setup.getTimeoutL0());
		
		click(".zstbtn-textIndentIncrease.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		click(".zstbtn-textIndentIncrease.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		click(".zstbtn-textIndentIncrease.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		
		long cellWidth = (Long) jq("@Row:eq(0) @Cell:eq(0)").fun("outerWidth");
		int realWidth = jq("@Row:eq(0) @Cell:eq(0) .zscelltxt-real").width();
		assertTrue(cellWidth < realWidth);
		
		click(".zstbtn-textIndentDecrease.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		click(".zstbtn-textIndentDecrease.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		click(".zstbtn-textIndentDecrease.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		click(".zstbtn-textIndentDecrease.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		
		assertTrue(jq("@Row:eq(0) @Cell:eq(0)").offsetLeft() < jq("@Row:eq(0) @Cell:eq(0) .zscelltxt-real").offsetLeft());
		
		click(".zschktbtn-protectSheet.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		assertEquals("disabled", jq(".zstbtn-textIndentDecrease.z-toolbarbutton").attr("disabled"));
		
		click(".zschktbtn-protectSheet.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		
		click(".zschktbtn-protectSheet.z-toolbarbutton");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-listitem:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		assertNotEquals("disabled", jq(".zstbtn-textIndentDecrease.z-toolbarbutton").attr("disabled"));
	}
}





