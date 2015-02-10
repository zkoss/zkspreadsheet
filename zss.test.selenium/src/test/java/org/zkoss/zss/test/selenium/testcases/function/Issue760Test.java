package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import java.util.Date;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue760Test extends ZSSTestCase {
	
	@Ignore("window sizing")
	@Test
	public void testZSS766() throws Exception {
		
	}
	
	@Test
	public void testZSS760() throws Exception {
		getTo("issue3/760-refresh.zul");
		
		JQuery a1 = jq(".zsrow:eq(0) .zscell:eq(0)");
		JQuery a3 = jq(".zsrow:eq(2) .zscell:eq(0)");
		
		assertTrue(a3.text().trim().equals("1"));
		
		click(jq(".z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(a1.text().trim().equals("2"));
		assertTrue(a3.text().trim().equals("3"));
		
		click(jq(".z-button:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(a1.text().trim().equals("2"));
		assertTrue(a3.text().trim().equals("3"));
		
		click(jq(".z-button:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(a1.text().trim().equals("3"));
		assertTrue(a3.text().trim().equals("3"));
		
		click(jq(".z-button:eq(3)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(a1.text().trim().equals("3"));
		assertTrue(a3.text().trim().equals("4"));
	}
	
	@Test
	public void testZSS762() throws Exception {
		getTo("issue3/762-sheetbar-content-menu.zul");
		
		rightClick(jq(".zssheettab:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL1());
		assertTrue(jq(".z-menupopup:visible").exists());
		
		rightClick(jq(".zssheettab:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL1());
		assertTrue(jq(".z-menupopup:visible").exists());
	}
}





