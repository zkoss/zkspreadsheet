package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import java.util.Date;

import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue810Test extends ZSSTestCase {
	
	@Test
	public void testZSS818() throws Exception {
		getTo("issue3/818-formula-update.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		click(ctrl.getCell(0, 5));
		waitUntilProcessEnd(Setup.getTimeoutL1());
		
		assertTrue(ctrl.getCell(2, 3).$n().getText().trim().equals("998"));
		assertTrue(ctrl.getCell(2, 5).$n().getText().trim().equals("998"));
	}
	
	@Test
	public void testZSS816() throws Exception {
		long duration = 15 * 1000;
		long t1 = new Date().getTime();
		getTo("issue3/818-formula-update.zul");
		long t2 = new Date().getTime();
		assertTrue(t2 - t1 < duration);
	}
}





