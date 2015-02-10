package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;


public class Issue890Test extends ZSSTestCase {
	
	@Ignore("download")
	@Test
	public void testZSS891() throws Exception {}
	
	@Ignore("download")
	@Test
	public void testZSS892() throws Exception {}
	
	@Ignore("download")
	@Test
	public void testZSS896() throws Exception {}
	
	@Ignore("duplicate with 864, 889")
	@Test
	public void testZSS898() throws Exception {}
	
	@Test
	public void testZSS899() throws Exception {
		getTo("/issue3/899-wrong-hyperlink.zul");
		focusSheet().getSheetCtrl();
		
		rightClick(jq(".zsrow:eq(4) .zscell:eq(4)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-hyperlink.z-menuitem"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(".z-window .z-combobox-input", "aaaa");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		// press OK
		click(jq(".z-window .z-button:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-error").exists());
		
		// omit step 7 ~ 11 due to selenium technical issue
	}
	
	
}





