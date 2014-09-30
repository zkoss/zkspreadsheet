package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue470Test extends ZSSTestCase {
	
	
	@Test
	public void testZSS472() throws Exception {
		getTo("/issue3/472-nameRange.zul");
		
		SpreadsheetWidget xlsxSpreadsheet = getSpreadsheet("$ss1");
		SheetCtrlWidget xlsxSheetCtrl = xlsxSpreadsheet.getSheetCtrl();
		EditorWidget xlsxEditor = xlsxSheetCtrl.getInlineEditor();
		xlsxSpreadsheet.focus();
		click(xlsxSheetCtrl.getCell("D4"));
		waitForTime(Setup.getTimeoutL0());
		xlsxEditor.toWebElement().sendKeys("2");
		waitForTime(Setup.getTimeoutL0());
		click(xlsxSheetCtrl.getCell("E2"));
		assertEquals("22", xlsxSheetCtrl.getCell("C2").getText());
		
		SpreadsheetWidget xlsSpreadsheet = getSpreadsheet("$ss2");
		SheetCtrlWidget xlsSheetCtrl = xlsSpreadsheet.getSheetCtrl();
		EditorWidget xlsEditor = xlsSheetCtrl.getInlineEditor();
		xlsSpreadsheet.focus();
		click(xlsSheetCtrl.getCell("D4"));
		waitForTime(Setup.getTimeoutL0());
		xlsEditor.toWebElement().sendKeys("2");
		waitForTime(Setup.getTimeoutL0());
		click(xlsSheetCtrl.getCell("E2"));
		assertEquals("22", xlsSheetCtrl.getCell("C2").getText());
	}
	
	@Ignore("Scroll")
	@Test
	public void testZSS475() throws Exception {
		
	}
	
	@Ignore("Export")
	@Test
	public void testZSS476() throws Exception {
		
	}
	
	@Ignore("Load")
	@Test
	public void testZSS477() throws Exception {
		getTo("/issue3/477-bordercolor.zul");
		
		SheetCtrlWidget xlsSheetCtrl = getSpreadsheet("$ss1").getSheetCtrl();
		SheetCtrlWidget xlsxSheetCtrl = getSpreadsheet("$ss2").getSheetCtrl();
		WebElement cellXlsA30 = xlsSheetCtrl.getCell("A30").$n();
		WebElement cellXlsxA30 = xlsxSheetCtrl.getCell("A30").$n();
		WebElement cellXlsA31 = xlsSheetCtrl.getCell("A31").$n();
		WebElement cellXlsxA31 = xlsxSheetCtrl.getCell("A31").$n();
		WebElement cellXlsE30 = xlsSheetCtrl.getCell("E30").$n();
		WebElement cellXlsxE30 = xlsxSheetCtrl.getCell("E30").$n();
		WebElement cellXlsE31 = xlsSheetCtrl.getCell("E31").$n();
		WebElement cellXlsxE31 = xlsxSheetCtrl.getCell("E31").$n();
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("1px", cellXlsA30.getCssValue("border-bottom-width"));
		assertEquals("1px", cellXlsA31.getCssValue("border-bottom-width"));
		assertEquals("1px", cellXlsA31.getCssValue("border-right-width"));
		assertEquals("1px", cellXlsxA30.getCssValue("border-bottom-width"));
		assertEquals("1px", cellXlsxA31.getCssValue("border-bottom-width"));
		assertEquals("1px", cellXlsxA31.getCssValue("border-right-width"));
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("rgba(204, 0, 0, 1)", cellXlsA30.getCssValue("border-bottom-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsA31.getCssValue("border-bottom-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsA31.getCssValue("border-right-color"));
		assertNotEquals("rgba(204, 0, 0, 1)", xlsSheetCtrl.getCell("B30").$n().getCssValue("border-left-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsxA30.getCssValue("border-bottom-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsxA31.getCssValue("border-bottom-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsxA31.getCssValue("border-right-color"));
		assertNotEquals("rgba(204, 0, 0, 1)", xlsxSheetCtrl.getCell("B30").$n().getCssValue("border-left-color"));
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("1px", cellXlsE30.getCssValue("border-bottom-width"));
		assertEquals("1px", cellXlsE31.getCssValue("border-bottom-width"));
		assertEquals("1px", cellXlsE31.getCssValue("border-right-width"));
		assertEquals("1px", cellXlsxE30.getCssValue("border-bottom-width"));
		assertEquals("1px", cellXlsxE31.getCssValue("border-bottom-width"));
		assertEquals("1px", cellXlsxE31.getCssValue("border-right-width"));
		click(jq("@button:eq(3)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("rgba(204, 0, 0, 1)", cellXlsE30.getCssValue("border-bottom-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsE31.getCssValue("border-bottom-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsE31.getCssValue("border-right-color"));
		assertNotEquals("rgba(204, 0, 0, 1)", xlsSheetCtrl.getCell("D30").$n().getCssValue("border-left-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsxE30.getCssValue("border-bottom-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsxE31.getCssValue("border-bottom-color"));
		assertEquals("rgba(204, 0, 0, 1)", cellXlsxE31.getCssValue("border-right-color"));
		assertNotEquals("rgba(204, 0, 0, 1)", xlsxSheetCtrl.getCell("D30").$n().getCssValue("border-left-color"));
	}
}
