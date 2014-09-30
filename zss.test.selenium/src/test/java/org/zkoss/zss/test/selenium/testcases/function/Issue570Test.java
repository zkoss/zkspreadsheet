package org.zkoss.zss.test.selenium.testcases.function;

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
	@Ignore("A bit complicated")
	@Test
	public void testZSS576() throws Exception {
		
	}
	
	@Ignore("Chart")
	@Test
	public void testZSS577() throws Exception {
		
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
