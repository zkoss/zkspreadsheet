package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue180Test extends ZSSTestCase {
	@Test
	public void testZSS182() throws Exception {
		getTo("/issue3/182-stop-edit-navigate.zul");
		
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		WebElement c4 = ctrl.getCell("C4").$n();
		WebElement body = jq("body").toWebElement();
		
		click(c4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(" aaa");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(Keys.ARROW_UP);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(c4.getText().trim(), "aaa");
		
		click(c4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(" bbb");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(Keys.ARROW_DOWN);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(c4.getText().trim(), "bbb");
		
		click(c4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(" ccc");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(Keys.ARROW_LEFT);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(c4.getText().trim(), "ccc");
		
		click(c4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(" ddd");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(Keys.ARROW_RIGHT);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(c4.getText().trim(), "ddd");
		
		click(c4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(" eee");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(Keys.HOME);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(c4.getText().trim(), "eee");
		
		click(c4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(" fff");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(Keys.END);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(c4.getText().trim(), "fff");
		
		click(c4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(" ggg");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(Keys.PAGE_UP);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(c4.getText().trim(), "ggg");
		
		click(c4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(" hhh");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		body.sendKeys(Keys.PAGE_DOWN);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(c4.getText().trim(), "hhh");
	}
}
