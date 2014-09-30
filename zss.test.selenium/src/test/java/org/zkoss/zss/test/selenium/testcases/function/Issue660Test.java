package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue660Test extends ZSSTestCase {
	@Ignore("It's OK")
	@Test
	public void testZSS660() throws Exception {
	}
	
	@Test
	public void testZSS664() throws Exception {
		getTo("/issue/664-page-up-down.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		EditorWidget inlineEditor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(sheetCtrl.getCell("C6"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		inlineEditor.toWebElement().sendKeys(Keys.PAGE_DOWN);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		inlineEditor.toWebElement().sendKeys(Keys.PAGE_UP);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq("@button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		JQuery result = jq("$result");
		waitUntil(1, ExpectedConditions.visibilityOf(result.toWebElement()));
		final String focusCell = result.text();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		inlineEditor.toWebElement().sendKeys(Keys.PAGE_DOWN);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		inlineEditor.toWebElement().sendKeys(Keys.PAGE_UP);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq("@button:eq(0)"));
		waitUntil(1, ExpectedConditions.visibilityOf(result.toWebElement()));
		assertEquals(focusCell, result.text());
	}
}
