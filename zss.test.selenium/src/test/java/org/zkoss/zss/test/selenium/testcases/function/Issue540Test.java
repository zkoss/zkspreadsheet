package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue540Test extends ZSSTestCase {
	@Ignore("Don't know how to assert")
	@Test
	public void testZSS540() throws Exception {
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS541() throws Exception {
		
	}
	
	@Test
	public void testZSS546() throws Exception {
		getTo("/issue3/546-left-key.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellC1 = sheetCtrl.getCell("C1");
		EditorWidget editor = sheetCtrl.getInlineEditor();
		
		spreadsheet.focus();
		click(cellC1);
		
		waitForTime(Setup.getTimeoutL0());
		
		editor.toWebElement().sendKeys(Keys.LEFT);
		
		waitForTime(Setup.getTimeoutL0());
		assertEquals("SEditbox", getActiveElement().getAttribute("zs.t"));
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS547() throws Exception {
		
	}
}
