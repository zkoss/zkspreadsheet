package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;


import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;

public class Issue330Test extends ZSSTestCase {
	
	@Test
	public void testZSS330() throws Exception {
		getTo("/issue3/330-hideMerge.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		CellWidget cellB2 = getSpreadsheet().getSheetCtrl().getCell("B2");
		assertEquals("Hello", cellB2.getText());
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("Hello", cellB2.getText());
	}
	
	@Ignore ("What?")
	@Test
	public void testZSS330_HideMerge() throws Exception {
		assertEquals("test test test", getSpreadsheet().getSheetCtrl().getCell("C1"));
	}
}
