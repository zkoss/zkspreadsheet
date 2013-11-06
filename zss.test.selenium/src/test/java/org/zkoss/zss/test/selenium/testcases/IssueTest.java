package org.zkoss.zss.test.selenium.testcases;

import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.ZSSTestcaseBase;

public class IssueTest extends ZSSTestcaseBase{
	@Test
	public void testZSS10() throws Exception{
		basename();
		
		getTo("/issue3/10-chart.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("btn0");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("btn1");
		
		
		click("@button:eq(2)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("btn2");
		
		click("@button:eq(3)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("btn3");
	}
	
	@Test
	public void testAnother() throws Exception{
		
	}

}
