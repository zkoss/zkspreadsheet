package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue700Test extends ZSSTestCase {
	@Test
	public void testZSS706() throws Exception {
		getTo("/issue3/706-click-between-style-panel.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		WebElement contextMenu = sheetCtrl.getContextMenu().$n();
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		rightClick(sheetCtrl.getCell(3, 2));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(contextMenu.isDisplayed());
		// click between two menu popups
		click(jq(ZSStyle.STYLEPANEL_MENU), 30, -10);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse(contextMenu.isDisplayed());
	}
}
