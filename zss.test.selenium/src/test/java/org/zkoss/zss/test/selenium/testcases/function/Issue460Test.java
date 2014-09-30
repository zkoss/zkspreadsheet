package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.ZSSTestCase;

public class Issue460Test extends ZSSTestCase {
	@Test
	public void testZSS463() throws Exception {
		getTo("/issue3/463-color-picker.zul");
		
		assertTrue(jq("@spreadsheet:eq(0) .zstbtn[title=\"Font Color\"]").exists());
		assertTrue(jq("@spreadsheet:eq(1) .zstbtn-fontColor").exists());
		assertTrue(jq("@spreadsheet:eq(2) .zstbtn[title=\"Font Color\"]").exists());
		assertTrue(jq("@spreadsheet:eq(3) .zstbtn[title=\"Font Color\"]").exists());
	}
	
	@Ignore("Too slow")
	@Test
	public void testZSS464() throws Exception {
	}
}
