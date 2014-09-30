package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.fail;

import org.junit.Test;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue340Test extends ZSSTestCase {
	
	@Test
	public void testZSS340() throws Exception {
		try {
			getTo("/issue3/340-npe-log.zul");
		} catch (RuntimeException exception) {
			fail("Should not trigger HTTP ERROR 500");
		}
	}
}
