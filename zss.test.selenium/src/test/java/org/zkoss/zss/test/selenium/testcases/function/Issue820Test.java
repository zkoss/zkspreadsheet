package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;


public class Issue820Test extends ZSSTestCase {
	
	@Test
	public void testZSS821() throws Exception {
		getTo("issue3/821-missing-chart-title.zul");
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL1());
		
		assertTrue(jq(".highcharts-title:eq(0)").text().equals("BB"));
		assertTrue(jq(".highcharts-title:eq(1)").text().equals("AA"));
		assertTrue(jq(".highcharts-title:eq(2)").text().equals("BB"));
		assertTrue(jq(".highcharts-title:eq(3)").text().equals("BB"));
		assertTrue(jq(".highcharts-title:eq(4)").text().equals("EBITDA Coverage (Pre/Post Tax)"));
		assertTrue(jq(".highcharts-title:eq(5)").text().equals("Current vs Quick Ratio"));
	}
	
	@Test
	public void testZSS822() throws Exception {
		getTo("issue3/822-axis-label-format.zul");
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(jq(".zswidget-cave:eq(0) text[text-anchor=\"end\"]:eq(1)").text().equals("54.56%"));
		assertTrue(jq(".zswidget-cave:eq(1) text[text-anchor=\"end\"]:eq(1)").text().equals("2.49%"));
		
		// skip export
	}
	
	@Ignore("vision")
	@Test
	public void testZSS828() throws Exception {
		
	}
}





