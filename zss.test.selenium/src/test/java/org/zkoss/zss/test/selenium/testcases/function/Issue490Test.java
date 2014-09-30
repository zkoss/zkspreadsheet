package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue490Test extends ZSSTestCase {

	@Test
	public void testZSS490() throws Exception {
		getTo("/issue3/490-externalReference.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS491() throws Exception {
		getTo("/issue3/491-changeColumnChart.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS492() throws Exception {
		getTo("/issue3/492-formula-rename-sheet.zul");
		
		click("@button:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		click("@button:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		assertEquals("abc", getSpreadsheet().getSheetCtrl().getCell("B1").getText());
	}
	
	@Test
	public void testZSS493() throws Exception {
		getTo("/issue3/493-deleteLastSheet.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Ignore()
	@Test
	public void testZSS494() throws Exception {
	}
	
	@Ignore()
	@Test
	public void testZSS496() throws Exception {
	}
	
	@Ignore("Chart")
	@Test
	public void testZSS499() throws Exception {
	}
}
