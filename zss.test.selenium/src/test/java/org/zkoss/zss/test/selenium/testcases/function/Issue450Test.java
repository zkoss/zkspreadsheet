package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue450Test extends ZSSTestCase {
	@Test
	public void testZSS450() throws Exception {
		getTo("/issue3/450-delete-row.zul");

		click(jq("@button:eq(3)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL1());
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("A140", getSpreadsheet().getSheetCtrl().getCell("A140").getText());
		assertEquals("D140", getSpreadsheet().getSheetCtrl().getCell("D140").getText());
	}
	
	@Ignore("Scroll")
	@Test
	public void testZSS451() throws Exception {
		
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS452() throws Exception {
	}
	
	@Ignore("Hyperlink")
	@Test
	public void testZSS454() throws Exception {
	}
	
	@Ignore("Hyperlink")
	@Test
	public void testZSS457() throws Exception {
	}
}
