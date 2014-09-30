package org.zkoss.zss.test.selenium.testcases.function;

import static org.hamcrest.number.OrderingComparison.lessThanOrEqualTo;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertThat;

import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue610Test extends ZSSTestCase {
	@Test
	public void testZSS616() throws Exception {
		getTo("/issue3/616-selection.zul");
		
		assertThat("Row 15 is visible.",
				getSpreadsheet().getSheetCtrl().getCell(14, 1).$n().getSize().getHeight(),
				lessThanOrEqualTo(1));
	}
	
	@Test
	public void testZSS617() throws Exception {
		getTo("/issue3/617-mask-blocking.zul");
		
		JQuery mask = jq(ZSStyle.MASK);
		
		assertFalse("A mask is blocking the sheet", mask.toWebElement().isDisplayed());
		click(jq("$xlsTab"));
		waitUntilProcessEnd(Setup.getTimeoutL1());
		assertFalse("A mask is blocking the sheet", mask.toWebElement().isDisplayed());
	}
}
