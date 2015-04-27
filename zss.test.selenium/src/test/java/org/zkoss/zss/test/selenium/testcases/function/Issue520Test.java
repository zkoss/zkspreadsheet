package org.zkoss.zss.test.selenium.testcases.function;


import static org.junit.Assert.*;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSSBy;

public class Issue520Test extends ZSSTestCase {
	
	@Test
	public void testZSS524() throws Exception {
		getTo("issue3/524-chart.zul");
		
		SheetFunction func = sheetFunction();
		click(".zssheettab:eq(1)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(".zssheettab:eq(2)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(".zssheettab:eq(3)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(".zssheettab:eq(5)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(".zssheettab:eq(6)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(".zssheettab:eq(7)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		AssertUtil.assertNoJAVAError();
	}
	
	@Ignore("vision")
	@Test
	public void testZSS525() throws Exception {
		
	}
	
	@Ignore("Bug is not fixed yet")
	@Test
	public void testZSS526_XLS() throws Exception {
	}
	
	@Ignore("Bug is not fixed yet")
	@Test
	public void testZSS526_XLST() throws Exception {
	}
	
	@Test
	public void testZSS528() throws Exception {
		getTo("/issue3/528-wrap.zul");
		SpreadsheetWidget ss = focusSheet();
		
		int height = jq("@Row:eq(0)").height();
		
		click(jq("@Row:eq(0) @Cell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq("@button:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(height > jq("@Row:eq(0)").height());
	}
}
