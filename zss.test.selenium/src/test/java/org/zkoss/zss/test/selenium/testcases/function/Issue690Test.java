package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue690Test extends ZSSTestCase {
	
	@Ignore("vision")
	@Test
	public void testZSS697() throws Exception {}
	
	@Test
	public void testZSS698() throws Exception {
		getTo("/issue3/698-validation-dialog.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
		// step 1
		rightClick(ctrl.getCell("A1"));
		waitForTime(Setup.getTimeoutL0());
		click(".zsmenuitem-dataValidation.z-menuitem");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "1");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "10");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-window.z-window-modal .z-tab:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(1) .z-textbox:eq(0)", "title1");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(1) .z-textbox:eq(1)", "message1");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-window.z-window-modal .z-tab:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(2) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(2) .z-textbox:eq(0)", "title2");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(2) .z-textbox:eq(1)", "message2");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		click("@button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		assertEquals("OK", jq("@textbox:eq(0)").val());
		
		
		// step 2
		click(ctrl.getCell("A1")); // gain focus first
		waitForTime(Setup.getTimeoutL0());
		rightClick(ctrl.getCell("A1"));
		waitForTime(Setup.getTimeoutL0());
		click(".zsmenuitem-dataValidation.z-menuitem");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(3)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", Keys.BACK_SPACE.toString());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "=A2:A4");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-window.z-window-modal .z-tab:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(1) @checkbox:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		
		click(".z-window.z-window-modal .z-tab:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(2) @checkbox:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		click("@button:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		assertEquals("OK", jq("@textbox:eq(0)").val());
	}
	
	@Test
	public void testZSS699() throws Exception {
		try {
			getTo("/issue3/699-crash.zul");
		} catch (RuntimeException exception) {
			fail("Should not trigger HTTP ERROR 500");
		}
	}
}
