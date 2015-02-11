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


public class Issue800Test extends ZSSTestCase {
	
	@Test
	public void testZSS809() throws Exception {
		getTo("issue3/809-literal-string-validation.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		JQuery b1 = jq(".zsrow:eq(0) .zscell:eq(0)");
		JQuery b2 = jq(".zsrow:eq(1) .zscell:eq(0)");
		JQuery b3 = jq(".zsrow:eq(2) .zscell:eq(0)");
		JQuery dropdown = jq(".zsdropdown");
		
		click(b1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(b1, dropdown, 7);
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("1"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals("2"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("3"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(3)").text().equals("4"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(4)").text().equals("5"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(5)").text().equals("6"));
		click(b1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(b2, dropdown, 7);
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().trim().equals(""));
		click(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(b3);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(b3, dropdown, 7);
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().trim().equals("a b c d e f"));
	}
}





