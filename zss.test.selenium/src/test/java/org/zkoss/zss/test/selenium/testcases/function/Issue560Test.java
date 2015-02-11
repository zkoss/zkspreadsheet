package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue560Test extends ZSSTestCase {
	@Test
	public void testZSS562() throws Exception {
		getTo("/issue3/562-bean.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellC3 = sheetCtrl.getCell("C3");
		EditorWidget editor = sheetCtrl.getInlineEditor();
		
		spreadsheet.focus();
		click(cellC3);
		
		waitForTime(Setup.getTimeoutL0());
		
		sendKeys(editor, "=");
		waitForTime(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys("assetsBean.list");
		waitForTime(Setup.getTimeoutL0());
		click(sheetCtrl.getCell("C4"));
		waitForTime(Setup.getTimeoutL1());
		assertEquals("#VALUE!", cellC3.getText());
	}
	
	@Ignore("vision")
	@Test
	public void testZSS568() throws Exception {
		
	}
	
	@Test
	public void testZSS569() throws Exception {
		getTo("/issue3/569-notify-bean-change.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellA1 = sheetCtrl.getCell("A1");
		CellWidget cellC2 = sheetCtrl.getCell("C2");
		CellWidget cellC3 = sheetCtrl.getCell("C3");
		CellWidget cellE2 = sheetCtrl.getCell("E2");
		WebElement tb1 = jq("$tb1").toWebElement();
		WebElement tb2 = jq("$tb2").toWebElement();
		
		tb1.sendKeys("AAA");
		click(cellE2);
		waitForTime(Setup.getTimeoutL0());
		assertEquals(tb1.getAttribute("value"), cellA1.getText());
		tb2.sendKeys("BBB");
		click(cellA1);
		waitForTime(Setup.getTimeoutL0());
		assertEquals(tb2.getAttribute("value"), cellC2.getText());
		assertEquals(tb2.getAttribute("value"), cellE2.getText());
		assertEquals("Value : " + tb2.getAttribute("value"), cellC3.getText());
	}
}
