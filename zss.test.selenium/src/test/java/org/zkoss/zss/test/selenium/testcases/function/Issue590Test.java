package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue590Test extends ZSSTestCase {

	@Test
	public void testZSS590() throws Exception {
		getTo("issue3/590-wrong-chart.zul");
		
		click(jq(".z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-button:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertEquals(jq(".zswidget-chart").length(), 1);
	}
}
