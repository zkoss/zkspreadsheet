package org.zkoss.zss.test.selenium;

import static org.junit.Assert.assertFalse;

import org.zkoss.zss.test.selenium.entity.JQuery;

public class AssertUtil {

	
	public static void assertNoJSError() {
		assertFalse("Should not show js error", JQuery.$(".z-error").exists());
	}
}
