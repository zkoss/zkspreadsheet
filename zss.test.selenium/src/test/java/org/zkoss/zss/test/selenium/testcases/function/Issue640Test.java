package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue640Test extends ZSSTestCase {
	

	@Test
	public void testZSS640() throws Exception {
		getTo("/issue3/640-freeze-update.zul");
		
		click(jq("@button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("Cell A1's value does not updated", "2",
				getSpreadsheet().getSheetCtrl().getCell("A1").getText());
		assertEquals("Cell B1's value does not updated", "2",
				getSpreadsheet().getSheetCtrl().getCell("B1").getText());
		assertEquals("Cell A2's value does not updated", "2",
				getSpreadsheet().getSheetCtrl().getCell("A2").getText());
		assertEquals("Cell B1's value does not updated", "2",
				getSpreadsheet().getSheetCtrl().getCell("B2").getText());
	}
}
