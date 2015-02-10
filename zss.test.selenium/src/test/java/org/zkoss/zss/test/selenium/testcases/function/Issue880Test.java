package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.interactions.Actions;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue880Test extends ZSSTestCase {
	
	@Ignore("SVG test")
	@Test
	public void testZSS882() throws Exception {
		
	}
	
	@Ignore("Resize window")
	@Test
	public void testZSS884() throws Exception {
		
	}
	
	@Ignore("SVG test")
	@Test
	public void testZSS885() throws Exception {

	}
	
	@Test
	public void testZSS886() throws Exception {
		getTo("/issue3/886-get-cell-js-err.zul");
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq("div.z-error").exists(), false);
	}
	
	@Test
	public void testZSS889() throws Exception {
		getTo("/issue3/889-unlock-cells-slow.zul");
		
		sheetFunction().gotoTab(4);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		SheetCtrlWidget ctrl = focusSheet().getSheetCtrl();
		click(ctrl.getCell(0, 0));
		waitUntilProcessEnd(Setup.getTimeoutL0());
				
		int height = jq(".zsselect").height();
		int width = jq(".zsselect").width();
		
		new Actions(driver()).dragAndDrop(
				ctrl.getCell(0, 0).toWebElement(), 
				ctrl.getCell(19, 7).toWebElement()).perform();
		
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		int height1 = jq(".zsselect").height();
		int width1 = jq(".zsselect").width();
		
		assertTrue(height * 20 <= height1 && width * 8 <= width1);
	}
}





