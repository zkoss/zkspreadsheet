package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue410Test extends ZSSTestCase {
	@Ignore()
	@Test
	public void testZSS412() throws Exception {
		getTo("/issue/412-overlap-undo.zul");
	}
	
	@Test
	public void testZSS414() throws Exception {
		getTo("/issue/414-client-paste-npe-original.zul");
		
		WebElement textbox = jq("@textbox").toWebElement();
		Actions action= new Actions(driver());
		action.moveToElement(textbox).click().sendKeys(Keys.chord(Keys.CONTROL, "a")).sendKeys(Keys.chord(Keys.CONTROL, "c")).perform();
		waitForTime(Setup.getTimeoutL0());
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		CellWidget cellD4 = sheetCtrl .getCell("D4");
		EditorWidget editor = sheetCtrl.getInlineEditor();
		spreadsheet.focus();
		click(cellD4);
		waitForTime(Setup.getTimeoutL0());
		editor.toWebElement().sendKeys(Keys.chord(Keys.CONTROL, "v"));
		waitForTime(Setup.getTimeoutL1());
		assertEquals("1", cellD4.getText());
		assertEquals("2", sheetCtrl.getCell("E4").getText());
		assertEquals("1", sheetCtrl.getCell("D5").getText());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS414_NPE() throws Exception {
		getTo("/issue/414-client-paste-npe.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS415() throws Exception {
		getTo("/issue/415-export-comment.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS418() throws Exception {
		
	}
	
	@Ignore("Export")
	@Test
	public void testZSS419() throws Exception {
		
	}
}
