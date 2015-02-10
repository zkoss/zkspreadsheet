package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue840Test extends ZSSTestCase {
	@Test
	public void testZSS849() throws Exception {
		getTo("issue3/849-cell-comment.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		JQuery jq = jq(".zscellcomment:not(.zscellcomment-hidden)");
		for(int i = jq.length() - 1; i >= 0; i--) {
			mouseMoveAt(jq.get(i), "1,1");
			waitUntilProcessEnd(Setup.getTimeoutL0());
			assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
			waitUntilProcessEnd(Setup.getTimeoutL0());
		}
		
		// 2. make sure removing then adding comment many times will show and hide E2 cell comment.
		click(jq(".z-button:nth(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		mouseMoveAt(ctrl.getCell(1, 4), "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
		click(jq(".z-button:nth(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		mouseMoveAt(ctrl.getCell(1, 4), "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(!jq(".zscellcomment-popup").css("display").equals("block"));
		
		// 3. merge cell D4:E6, check merged cell will show comment in non-select, select and copy mode.
		CellWidget c1 = ctrl.getCell(3, 3);
		CellWidget c2 = ctrl.getCell(7, 3);
		CellWidget c3 = ctrl.getCell(11, 3);
		
		dragAndDrop(c1.toWebElement(), ctrl.getCell(5, 4).toWebElement());
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetFunction().mergeToggle();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetFunction().copy();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(c2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetFunction().paste();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(c1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetFunction().copy();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(c3);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sheetFunction().paste();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(c2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(c1, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
		
		mouseMoveAt(c2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
		
		mouseMoveAt(c3, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
		
		// 4. freeze sheet and then do step 1 and 2 again.
		// 5. switch sheets, make sure all functions above will work as well.
		click(jq(".z-button:nth(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		sheetFunction().gotoTab(2);
		waitUntilProcessEnd(Setup.getTimeoutL1());
		sheetFunction().gotoTab(1);
		waitUntilProcessEnd(Setup.getTimeoutL1());
		
		jq = jq(".zscellcomment:not(.zscellcomment-hidden)");
		for(int i = jq.length() - 1; i >= 0; i--) {
			mouseMoveAt(jq.get(i), "1,1");
			waitUntilProcessEnd(Setup.getTimeoutL0());
			assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
			waitUntilProcessEnd(Setup.getTimeoutL0());
		}
		
		click(jq(".z-button:nth(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		mouseMoveAt(ctrl.getCell(1, 4), "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
		click(jq(".z-button:nth(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		mouseMoveAt(ctrl.getCell(1, 4), "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("block"));
	}
}





