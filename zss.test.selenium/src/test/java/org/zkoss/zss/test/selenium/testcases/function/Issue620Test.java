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

public class Issue620Test extends ZSSTestCase {
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS621() throws Exception {
	}
	
	@Test
	public void testZSS622() throws Exception {
		getTo("/issue3/622-runtime.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellE7 = sheetCtrl.getCell("E7");
		EditorWidget editor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		click(cellE7);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(editor, "AAA");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(ZSStyle.FONT_BOLD_BTN));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("Font weight should be bold in E7", "bold", cellE7.$n("real").getCssValue("font-weight"));		
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS623() throws Exception {
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS624() throws Exception {
	}
}
