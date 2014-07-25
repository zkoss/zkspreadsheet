package org.zkoss.zss.test.selenium.testcases;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCaseBase;

public class IssueTest extends ZSSTestCaseBase {
	@Test
	public void testZSS10() throws Exception{
		basename();
		
		getTo("/issue3/10-chart.zul");
		waitForTime(Setup.getTimeoutL2());
		captureOrAssert("loadpage");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL2());//for zss render in browser
		captureOrAssert("btn0");
		
		click("@button:eq(1)");
		waitForTime(Setup.getTimeoutL2());//for zss render in browser
		captureOrAssert("btn1");
		
		
		click("@button:eq(2)");
		waitForTime(Setup.getTimeoutL2());//for zss render in browser
		captureOrAssert("btn2");
		
		click("@button:eq(3)");
		waitForTime(Setup.getTimeoutL2());//for zss render in browser
		captureOrAssert("btn3");
	}
	
	@Test
	public void testZSS202() throws Exception{
		basename();
		
		getTo("/issue3/202-chartTitle.zul");
		waitForTime(Setup.getTimeoutL1());
		captureOrAssert("loadpage");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());//for zss render in browser
	}
	
	@Test
	public void testZSS219() throws Exception{
		basename();
		
		getTo("/issue3/219-focusto.zul");
		waitForTime(Setup.getTimeoutL1());
		
		for(int i = 0; i < 9; i++) {
			click("@button:eq("+ i +")");
			waitForTime(Setup.getTimeoutL2());//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	@Test
	public void testZSS219_2() throws Exception{
		basename();
		
		getTo("/issue3/219-focusto.zul");
		waitForTime(Setup.getTimeoutL1());
		
		click("@button:eq(8)");
		captureOrAssert("freeze-first");
		
		for(int i = 0; i < 8; i++) {
			click("@button:eq("+ i +")");
			waitForTime(Setup.getTimeoutL3());//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	@Test
	public void testZSS242() throws Exception {
		basename();
		
		getTo("/issue3/242-maxrow.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		for(int i = 0; i < 17; i++) {
			click("@button:eq("+ i +")");
			waitForTime(Setup.getTimeoutL3());//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	@Test
	public void testZSS250() throws Exception{
		basename();
		
		getTo("/issue3/250-deleteSheet.zul");
		waitForTime(Setup.getTimeoutL1());
		captureOrAssert("loadpage");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());//for zss render in browser
		captureOrAssert("deleteFirstSheet");
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());//for zss render in browser
		captureOrAssert("deleteSecondSheet");
	}
	
	@Test
	public void testZSS256() throws Exception{
		basename();
		
		getTo("/issue3/256-rowcolumn.zul");
		waitForTime(Setup.getTimeoutL1());
		captureOrAssert("loadpage");	
	}
	
	@Test
	public void testZSS256XLS() throws Exception{
		basename();
		
		getTo("/issue3/256-rowcolumn-xls.zul");
		waitForTime(Setup.getTimeoutL1());
		captureOrAssert("loadpage");	
	}
	
	@Test
	public void testZSS282() throws Exception{
		basename();
		
		getTo("/issue3/282-paste-merge.zul");
		waitForTime(Setup.getTimeoutL1());
		captureOrAssert("loadpage");
		
		for(int i = 0; i < 7; i++) {
			click("@button:eq("+ i +")");
			waitForTime(Setup.getTimeoutL0());//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	
	@Ignore
	@Test
	public void testAnother() throws Exception{
		basename();
		
		getTo("/issue3/401-cutMerged.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		sendZSSKeys(Keys.chord(Keys.CONTROL, "z"));
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");		
//		click("@button:eq(1)");
//		waitForTime(500);//for zss render in browser
//		captureOrAssert("step2");
	}

}
