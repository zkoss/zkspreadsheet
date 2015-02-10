package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue740Test extends ZSSTestCase {
	
	@Ignore("vision, need export")
	@Test
	public void testZSS742() throws Exception {
		
	}
	
	@Test
	public void testZSS744() throws Exception {
		getTo("issue3/744-toolbarbutton-stopedit.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		CellWidget cell = ctrl.getCell("A1"); 
		click(cell);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(jq("body"), " AAA");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".zstbtn-fontBold:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(!eval("return jq('@spreadsheet').zk.$().sheetCtrl.state").equals("6"));
		assertTrue(jq(".zsrow:eq(0) .zscell:eq(0) .zscelltxt-real").css("font-weight").trim().equals("bold"));
	}
	
	@Ignore("Need export")
	@Test
	public void testZSS746() throws Exception {
		
	}
	
	@Ignore("Need export")
	@Test
	public void testZSS748() throws Exception {
		
	}
	
	
}
