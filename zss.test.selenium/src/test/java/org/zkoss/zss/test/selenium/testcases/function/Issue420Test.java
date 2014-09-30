package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue420Test extends ZSSTestCase {
	@Ignore("Scroll")
	@Test
	public void testZSS420() throws Exception {
		
	}
	
	@Ignore("Scroll")
	@Test
	public void testZSS421() throws Exception {
		
	}
	
	@Test
	public void testZSS423() throws Exception {
		getTo("/issue/423-bug-freeze.zul");
		
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS428() throws Exception {
		getTo("/issue3/428-chart-xlsx.zul");
		
		AssertUtil.assertNoJSError();
	}
}
