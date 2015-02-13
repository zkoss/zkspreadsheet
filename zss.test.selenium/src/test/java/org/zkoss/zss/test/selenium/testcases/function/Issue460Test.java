package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue460Test extends ZSSTestCase {
	@Test
	public void testZSS463() throws Exception {
		getTo("/issue3/463-color-picker.zul");
		
		assertTrue(jq("@spreadsheet:eq(0) .zstbtn[title=\"Font Color\"]").exists());
		assertTrue(jq("@spreadsheet:eq(1) .zstbtn-fontColor").exists());
		assertTrue(jq("@spreadsheet:eq(2) .zstbtn[title=\"Font Color\"]").exists());
		assertTrue(jq("@spreadsheet:eq(3) .zstbtn[title=\"Font Color\"]").exists());
	}
	
	@Test
	public void testZSS464() throws Exception {
		getTo("issue3/464-manystyle-xlsx.zul");
		
		click(".z-button:eq(0)");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		eval("return jq('.z-loading-indicator:visible').size()");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(".z-button:eq(1)");
		waitUntilProcessEnd(Setup.getTimeoutL1());
		
		assertEquals(eval("return jq(jq('@spreadsheet').zk.$().sheetCtrl.getCell(99, 49).comp).css('background-color')"), "rgba(0, 0, 0, 0)");
		assertEquals(eval("return jq(jq('@spreadsheet').zk.$().sheetCtrl.getCell(99, 49).comp).css('font-size')"), "16px");
	}
}
