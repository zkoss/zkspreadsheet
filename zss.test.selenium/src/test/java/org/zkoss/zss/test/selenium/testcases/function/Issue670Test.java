package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue670Test extends ZSSTestCase {
	@Ignore("Export")
	@Test
	public void testZSS670() throws Exception {
	}
	
	@Test
	public void testZSS674() throws Exception {	
		getTo("/issue/674-save-focus-cell.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellA1 = sheetCtrl.getCell("A1");
		EditorWidget editor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(cellA1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(editor, "1111");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(ZSStyle.SAVE_BOOK_BTN));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		driver().navigate().refresh();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		final String textA1 = cellA1.getText();
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(cellA1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys("XX");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		editor.toWebElement().clear();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(ZSStyle.SAVE_BOOK_BTN));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("A1 text not saved", "1111", textA1);
	}
}
