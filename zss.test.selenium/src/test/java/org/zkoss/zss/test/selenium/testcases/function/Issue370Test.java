package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue370Test extends ZSSTestCase {
	
	@Test
	public void testZSS376() throws Exception {
		try {
			getTo("/issue3/376-chart-xlsx.zul");
		} catch (RuntimeException exception) {
			fail("Should not trigger HTTP ERROR 500");
		}
	}
	
	@Ignore("What")
	@Test
	public void testZSS379() throws Exception {
	}
}
