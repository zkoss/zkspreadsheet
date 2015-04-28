package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue430Test extends ZSSTestCase {
	@Ignore("Export")
	@Test
	public void testZSS430() throws Exception {

	}
	
	@Test
	public void testZSS433() throws Exception {
		getTo("/issue3/433-column-width.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
	
	@Ignore("Javascript Error")
	@Test
	public void testZSS433_1() throws Exception {
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS434() throws Exception {
		
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS434_1() throws Exception {
		
	}
	
	@Test
	public void testZSS435() throws Exception {
		getTo("/issue3/435-merge-border.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("1px", getSpreadsheet().getSheetCtrl()
				.getCell("C1").$n().getCssValue("border-bottom-width"));
	}
	
	@Test
	public void testZSS436() throws Exception {
		getTo("/issue3/436-wrongview.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL1());
		AssertUtil.assertNoJSError();
	}
	
	@Ignore("Browser Zoom")
	@Test
	public void testZSS438() throws Exception {
		
	}
}
