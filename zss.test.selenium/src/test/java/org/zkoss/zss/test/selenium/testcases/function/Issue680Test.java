package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue680Test extends ZSSTestCase {
	@Test
	public void testZSS681() throws Exception {	
		getTo("/issue3/681-stylepanel-floatup.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		EditorWidget editor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		rightClick(sheetCtrl.getCell("C4"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(editor, "\t");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse("Context Menu is not closed",
				sheetCtrl.getContextMenu().$n().isDisplayed());
	}
	
	@Test
	public void testZSS683() throws Exception {
		getTo("/issue3/683-infinite-error-msg-formula.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		EditorWidget inlineEditor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(sheetCtrl.getCell("C4"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(inlineEditor, "=");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(inlineEditor,"SUM(A1, B1");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(sheetCtrl.getCell("C5"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue("Should show the warning", jq(".z-messagebox-window").exists());
		click(jq(".z-window-icon.z-window-close"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse("Should not show the warning", jq(".z-messagebox-window").exists());
		inlineEditor.toWebElement().clear();
		click(sheetCtrl.getCell("C6"));
		EditorWidget formulabarEditor = sheetCtrl.getFormulabarEditor();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(formulabarEditor);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(formulabarEditor, "=");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(formulabarEditor,"SUM(A1, B1");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(sheetCtrl.getCell("C7"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue("Should show the warning", jq(".z-messagebox-window").exists());
		click(jq(".z-window-icon.z-window-close"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse("Should not show the warning", jq(".z-messagebox-window").exists());
	}
	
	@Test
	public void testZSS688() throws Exception {
		getTo("/issue3/688-clonesheet.zul");
		
		final JQuery sheetData = jq(ZSStyle.DATA);
		final String data = sheetData.text();
		click(jq("@button:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq("@button:eq(3)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(data, sheetData.text());
	}
	
	@Test
	public void testZSS689() throws Exception {
		getTo("/issue3/689-insert-merged-cell-ref.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		EditorWidget editor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(sheetCtrl.getCell("A1"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(editor, "=");
		waitUntilProcessEnd(Setup.getTimeoutL0());
	
		click(sheetCtrl.getCell("D6"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("Cell reference is wrong when focus on merged cell", "=C4", editor.getValue());
	}
}
