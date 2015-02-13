package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;


public class Issue860Test extends ZSSTestCase {
	@Test
	public void testZSS861() throws Exception {
		getTo("issue3/861-merge-focus.zul");
		
		click(jq(".zscell").get(0));
		
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		int sleft = jq(".zsselect").offsetLeft();
		int stop = jq(".zsselect").offsetTop();
		
		int fleft = jq(".zsfocmark").offsetLeft();
		int ftop = jq(".zsfocmark").offsetTop();
		
		assertTrue(sleft == fleft && stop == ftop);
	}
	
	@Ignore("svg")
	@Test
	public void testZSS862() throws Exception {}
	
	@Test
	public void testZSS863() throws Exception {
		getTo("issue3/863-copy-focus-mark.zul");
		
		click(jq(".zscell").get(0));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zstbtn-copy.zstbtn.z-toolbarbutton"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq("@button"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zshighlight").css("display").equals("none"));
	}
	
	@Test
	public void testZSS864() throws Exception {
		getTo("issue3/864-select-comment.zul");
		
		focusSheet();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		rightClick(jq(".zsrow:eq(4) .zscell:eq(4)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-insertComment"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		this.sendKeys(".z-window iframe", "test");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq("@button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(
			jq(".zsrow:eq(4) .zscell:eq(4) .zscellcomment").css("border-color").equals(
			"rgb(255, 0, 0) rgb(255, 0, 0) rgba(0, 0, 0, 0) rgba(0, 0, 0, 0)"));
		
		mouseMoveAt(jq(".zsrow:eq(4) .zscell:eq(4)"), "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
		
		dragAndDrop(jq(".zsrow:eq(2) .zscell:eq(1)").get(0),
				jq(".zsrow:eq(6) .zscell:eq(6)").get(0));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		mouseMoveAt(jq(".zsrow:eq(4) .zscell:eq(4)"), "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
	}
	
	@Ignore("vision issue")
	@Test
	public void testZSS865() throws Exception {}
	
	@Test
	public void testZSS866() throws Exception {
		getTo("issue3/866-validation-percentage.zul");
		
		focusSheet();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zscell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdropdown"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdv-item").get(1));
		AssertUtil.assertNoJSError();
		
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsrow:eq(1) .zscell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdropdown"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdv-item").get(0));
		AssertUtil.assertNoJSError();
		
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsrow:eq(2) .zscell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdropdown"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdv-item").get(0));
		AssertUtil.assertNoJSError();

		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsrow:eq(3) .zscell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdropdown"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdv-item").get(1));
		AssertUtil.assertNoJSError();
		
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsrow:eq(4) .zscell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdropdown"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsdv-item").get(2));
		AssertUtil.assertNoJSError();
	}
}





