package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue710Test extends ZSSTestCase {

	@Test
	public void testZSS710() throws Exception {
		getTo("issue3/710-sheettab-keycode-del.zul");
		
		doubleClick(jq(".zssheettab:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		JQuery textbox = jq(".zssheettab-rename-textbox:visible:eq(0)");
		textbox.toWebElement().sendKeys(Keys.DELETE);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zssheettab-rename-textbox:visible").val().trim(), "");
		textbox.toWebElement().sendKeys(Keys.ESCAPE);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zssheettab:eq(0)").text().trim(), "Sheet1");
		
		doubleClick(jq(".zssheettab:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		textbox.toWebElement().sendKeys("Sheetn");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		textbox.toWebElement().sendKeys(Keys.ENTER);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zssheettab:eq(0)").text().trim(), "Sheetn");
	}
}
