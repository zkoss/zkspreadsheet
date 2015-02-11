package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import org.junit.Test;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue830Test extends ZSSTestCase {
	
	@Test
	public void testZSS833() throws Exception {
		getTo("issue3/833-relative-validation.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		JQuery a1 = jq(".zsrow:eq(0) .zscell:eq(0)");
		JQuery b1 = jq(".zsrow:eq(0) .zscell:eq(1)");
		JQuery c1 = jq(".zsrow:eq(0) .zscell:eq(2)");
		JQuery f1 = jq(".zsrow:eq(0) .zscell:eq(5)");
		JQuery g1 = jq(".zsrow:eq(0) .zscell:eq(6)");
		JQuery h1 = jq(".zsrow:eq(0) .zscell:eq(7)");
		JQuery dropdown = jq(".zsdropdown");
		
		// 1. Select A1, B1, C1, F1, G1, H1; you should see a dropdown button by them.
		click(a1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(a1, dropdown, 7);
		
		click(b1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(b1, dropdown, 7);
		
		click(c1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(c1, dropdown, 7);
		
		click(f1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(f1, dropdown, 7);
		
		click(g1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(g1, dropdown, 7);
		
		click(h1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(h1, dropdown, 7);
		
		click(a1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("a4"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals("a5"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("a6"));
		click(a1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(b1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("b4"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals("b5"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("b6"));
		click(b1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(c1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("c4"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals("c5"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("c6"));
		click(c1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(f1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("f4"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals("f5"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("f6"));
		click(f1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(g1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals(""));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals(""));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals(""));
		click(g1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(h1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals(""));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals(""));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals(""));
		click(h1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		CellWidget widget = ctrl.getCell(4, 1); 
		click(widget);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		// selenium can't enter first character
		sendKeys(jq("body"), " b7");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(b1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("b4"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().trim().equals("b7"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("b6"));
	}
	
	
	
	@Test
	public void testZSS834() throws Exception {
		getTo("issue3/834-nameRange-validation.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		JQuery b1 = jq(".zsrow:eq(0) .zscell:eq(1)");
		JQuery b2 = jq(".zsrow:eq(1) .zscell:eq(1)");
		JQuery dropdown = jq(".zsdropdown");
		
		click(b1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(b1, dropdown, 7);
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("Item 1"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals("Item 2"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("Item 3"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(3)").text().equals("Item 4"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(4)").text().equals("Item 5"));
		click(jq(".zscellpopup:visible .zsdv-item:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(b2, dropdown, 7);
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("11"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals("21"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("31"));
		click(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(b1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zscellpopup:visible .zsdv-item:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetItemUtil().assertHorizontalCloseTo(b2, dropdown, 7);
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(0)").text().equals("10"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(1)").text().equals("20"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(2)").text().equals("30"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(3)").text().equals("40"));
		assertTrue(jq(".zscellpopup:visible .zsdv-item:eq(4)").text().equals("50"));
	}

	@Test
	public void testZSS837() throws Exception {
		getTo("issue3/837-delete-validation.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		click(ctrl.getCell(0, 1));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(jq(".zsdropdown").exists());
		
		click(jq(".z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(ctrl.getCell(0, 1));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(!jq(".zsdropdown").exists());
	}
}





