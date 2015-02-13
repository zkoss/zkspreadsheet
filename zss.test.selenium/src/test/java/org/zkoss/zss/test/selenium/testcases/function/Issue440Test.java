package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue440Test extends ZSSTestCase {
	@Ignore("browser zoom in")
	@Test
	public void testZSS440() throws Exception {
		
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS442() throws Exception {
		
	}
	
	@Test
	public void testZSS443() throws Exception {
		getTo("issue3/443-column-width.zul");
		
		click(".z-button:eq(0)");
		waitUntilProcessEnd(Setup.getTimeoutL0());

		click(".z-button:eq(1)");
		waitUntilProcessEnd(Setup.getTimeoutL0());

		click(".z-button:eq(2)");
		waitUntilProcessEnd(Setup.getTimeoutL0());

		click(".z-button:eq(3)");
		waitUntilProcessEnd(Setup.getTimeoutL0());

		click(".z-button:eq(4)");
		waitUntilProcessEnd(Setup.getTimeoutL0());

		click(".z-button:eq(5)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		AssertUtil.assertNoJSError();
	}
	
	@Ignore("A bit complicated")
	@Test
	public void testZSS448() throws Exception {
	}
}
