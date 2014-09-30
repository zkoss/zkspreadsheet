package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue600Test extends ZSSTestCase {
	
	@Ignore("Rewrite condition")
	@Test
	public void testZSS608() throws Exception {
		getTo("/issue3/608-pasteSingleToMerge.zul");

		click(jq("@button:eq(0)"));
		CellWidget cellA1 = getSpreadsheet().getSheetCtrl().getCell("A1");
		//waitUntil(1, ExpectedConditions.textToBePresentInElement(cellA1.$n(), "100"));
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("100", getSpreadsheet().getSheetCtrl().getCell(2, 2).getText());
	}
	
	@Test
	public void testZSS609() throws Exception {
		getTo("/issue/609-empty-sheet-name.zul");
		
		WebElement sheet1 = jq(ZSStyle.SHEET_TAB + ":eq(0)").toWebElement();
		doubleClick(sheet1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheet1.findElement(By.className("z-textbox")).clear();
		click(jq(ZSStyle.SHEET_TAB + ":eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue("Should show the warning", jq(".z-messagebox-window").exists());
	}
}
