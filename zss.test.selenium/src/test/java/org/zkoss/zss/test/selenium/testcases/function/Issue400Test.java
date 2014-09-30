package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;

public class Issue400Test extends ZSSTestCase {
	
	@Test
	public void testZSS401() throws Exception {
		getTo("/issue3/401-cutMerged.zul");
		
		CellWidget cellA1 = getSpreadsheet().getSheetCtrl().getCell("A1");
		final String text = cellA1.getText();
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals(text, cellA1.getText());
	}
	
	@Test
	public void testZSS401_Merge() throws Exception {
		getTo("/issue3/401-mergeAnotherSheet.zul");
		
		CellWidget cellA1 = getSpreadsheet().getSheetCtrl().getCell("A1");
		final String text = cellA1.getText();
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		assertTrue(cellA1.isMerged());
		assertEquals(text, cellA1.getText());
	}
	
	@Test
	public void testZSS408() throws Exception {
		getTo("/issue/408-save-autofilter.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
		waitForTime(Setup.getTimeoutL0());
	}

	@Test
	public void testZSS409() throws Exception {
		getTo("/issue/409-font-color.zul");
		
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		AssertUtil.assertNoJSError();
	}
}
