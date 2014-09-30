package org.zkoss.zss.test.selenium.testcases.function;


import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.ZSSBy;

public class Issue520Test extends ZSSTestCase {
	
	@Ignore("Chart")
	@Test
	public void testZSS524() throws Exception {
	}
	
	@Ignore("Chart")
	@Test
	public void testZSS525() throws Exception {
	}
	
	@Ignore("Bug is not fixed yet")
	@Test
	public void testZSS526_XLS() throws Exception {
	}
	
	@Ignore("Bug is not fixed yet")
	@Test
	public void testZSS526_XLST() throws Exception {
	}
	
	@Test
	public void testZSS528() throws Exception {
		getTo("/issue3/528-wrap.zul");
		
		click(jq("@button:eq(1)"));
		waitUntil(2, ExpectedConditions.textToBePresentInElement(ZSSBy.javascript("jq($msgArea)[0]"), "resize"));
	}
}
