package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import java.util.Date;

import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue750Test extends ZSSTestCase {
	
	@Test
	public void testZSS756() throws Exception {
		getTo("issue3/756-click-insert-function.zul");
		
		click(jq(".zsformulabar-insertbtn"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(jq(".z-window .z-textbox:eq(0)"), "series");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		int height = jq(".z-listcell-content:visible:eq(0)").height();
		click(jq(".z-window .z-listbox:eq(0)"), 10, height * 5);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS755() throws Exception {
		getTo("issue3/755-firefox-minus-key.zul");
		
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget xlsxEditor = ctrl.getInlineEditor();
		
		CellWidget cell = ctrl.getCell(1, 1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(cell);
		doubleClick(cell);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		xlsxEditor.toWebElement().sendKeys("-");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(ctrl.getCell(0, 0));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(cell.$n().getText().trim().equals("-"));
	}
}





