package org.zkoss.zss.test.selenium.testcases.function;

import static org.hamcrest.number.OrderingComparison.lessThanOrEqualTo;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertThat;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;

public class Issue580Test extends ZSSTestCase {

	@Test
	public void testZSS584() throws Exception {
		getTo("/issue3/584-wrong-color.zul");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		
		SheetCtrlWidget sheetCtrl = getSpreadsheet().getSheetCtrl();
		assertEquals("rgba(255, 0, 0, 1)", sheetCtrl.getCell("E7").$n()
				.getCssValue("background-color"));	
	}
	
	@Ignore("vision")
	@Test
	public void testZSS587() throws Exception {
		
	}
	
	@Test
	public void testZSS589() throws Exception {
		getTo("/issue3/589-hidden-row.zul");
		
		WebElement cellA15 = getSpreadsheet().getSheetCtrl().getCell("A15").$n();
		assertThat("Row 15 is visible.", cellA15.getSize().getHeight(),
				lessThanOrEqualTo(1));
		click(jq("$xlsTab"));
		waitUntilProcessEnd(Setup.getTimeoutL1());
		assertThat("Row 15 is visible.", cellA15.getSize().getHeight(),
				lessThanOrEqualTo(1));
	}
}
