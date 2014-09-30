package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.interactions.Actions;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue500Test extends ZSSTestCase {
	
	
	@Test
	public void testZSS501() throws Exception {
		getTo("/issue3/501-formulaBarCheck.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellA1 = sheetCtrl.getCell("A1");
		CellWidget cellC4 = sheetCtrl.getCell("C4");
		EditorWidget formulabarEditor = sheetCtrl.getFormulabarEditor();
		spreadsheet.focus();
		click(cellC4);
		waitForTime(Setup.getTimeoutL0());
		click(formulabarEditor);
		waitForTime(Setup.getTimeoutL0());
		sendKeys(formulabarEditor, "AAA");
		click(jq(ZSStyle.FORMULABAR_OK_BTN));
		waitForTime(Setup.getTimeoutL0());
		click(cellA1);
		waitForTime(Setup.getTimeoutL0());
		click(formulabarEditor);
		waitForTime(Setup.getTimeoutL0());
		sendKeys(formulabarEditor, "BBB");
		waitForTime(Setup.getTimeoutL0());
		click(cellC4);
		assertEquals("A1 is uneditable", "BBB", cellA1.getText());
	}
	
	@Test
	public void testZSS506() throws Exception {
		getTo("/issue3/506-deleteChartData.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		spreadsheet.focus();
		JQuery chart = jq(ZSStyle.WIDGET);
		click(chart);
		waitForTime(Setup.getTimeoutL0());
		Actions action= new Actions(driver());
		action.moveToElement(chart.toWebElement()).click().sendKeys(Keys.DELETE).perform();
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS509() throws Exception {
		getTo("/issue3/509-columnLocked.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellC3 = sheetCtrl.getCell("C3");
		CellWidget cellD3 = sheetCtrl.getCell("D3");
		CellWidget cellG3 = sheetCtrl.getCell("G3");
		EditorWidget editor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		click(cellC3);
		final String textC3 = cellC3.getText();
		waitForTime(Setup.getTimeoutL0());
		sendKeys(editor, "AAA");
		waitForTime(Setup.getTimeoutL0());
		click(cellD3);
		waitForTime(Setup.getTimeoutL0());
		assertEquals("Locked cell should not be editable", textC3, cellC3.getText());
		final String textD3 = cellD3.getText();
		waitForTime(Setup.getTimeoutL0());
		sendKeys(editor, "AAA");
		waitForTime(Setup.getTimeoutL0());
		click(cellG3);
		waitForTime(Setup.getTimeoutL0());
		assertEquals("Locked cell should not be editable", textD3, cellD3.getText());
		sendKeys(editor, "AAA");
		waitForTime(Setup.getTimeoutL0());
		click(cellC3);
		waitForTime(Setup.getTimeoutL0());
		assertEquals("Locked cell should be editable", "AAA", cellG3.getText());
	}
}
