package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.*;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.EditorWidget;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;


public class Issue980Test extends ZSSTestCase {
	
	@Test
	public void testZSS981() throws Exception {
		getTo("/issue3/981-validation-verify.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		EditorWidget editor = ctrl.getInlineEditor();
		
		// 1
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
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		// 2
		rightClick(ctrl.getCell("A2"));
		waitForTime(Setup.getTimeoutL0());
		click(".zsmenuitem-dataValidation.z-menuitem");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "1.1");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "=10.1");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		// 3
		rightClick(ctrl.getCell("A3"));
		waitForTime(Setup.getTimeoutL0());
		click(".zsmenuitem-dataValidation.z-menuitem");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(4)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "1/30/2015");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "1/31");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		// 4
		rightClick(ctrl.getCell("A4"));
		waitForTime(Setup.getTimeoutL0());
		click(".zsmenuitem-dataValidation.z-menuitem");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(5)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "14:20");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "3:20 PM");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		// 5
		rightClick(ctrl.getCell("A5"));
		waitForTime(Setup.getTimeoutL0());
		click(".zsmenuitem-dataValidation.z-menuitem");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(6)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "1");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "1000");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		
		click(ctrl.getCell("A1"));
		waitForTime(Setup.getTimeoutL0());
		sendKeys(editor, "5");
		waitForTime(Setup.getTimeoutL0());
		click(ctrl.getCell("A2"));
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-error").length() == 0);
		click(ctrl.getCell("A1"));
		waitForTime(Setup.getTimeoutL0());
		sendKeys(editor, "11");
		waitForTime(Setup.getTimeoutL0());
		click(ctrl.getCell("A2"));
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-error").length() > 0);
		click("@button:contains(Cancel)");
		
		click(ctrl.getCell("A2"));
		waitForTime(Setup.getTimeoutL0());
		sendKeys(editor, "5.1");
		waitForTime(Setup.getTimeoutL0());
		click(ctrl.getCell("A3"));
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-error").length() == 0);
		click(ctrl.getCell("A2"));
		waitForTime(Setup.getTimeoutL0());
		sendKeys(editor, "11");
		waitForTime(Setup.getTimeoutL0());
		click(ctrl.getCell("A3"));
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-error").length() > 0);
		click("@button:contains(Cancel)");
		
		// test constraint
		rightClick(ctrl.getCell("B1"));
		waitForTime(Setup.getTimeoutL0());
		click(".zsmenuitem-dataValidation.z-menuitem");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "1.1");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "10.1");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-exclamation").length() > 0);
		click(".z-window-highlighted @button:contains(OK)");

		String backSpace = Keys.BACK_SPACE.toString();
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", backSpace + backSpace + backSpace + backSpace);
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-exclamation").length() > 0);
		click(".z-window-highlighted @button:contains(OK)");
		
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", backSpace + backSpace + backSpace);
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "2");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "1");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-exclamation").length() > 0);
		click(".z-window-highlighted @button:contains(OK)");
		
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(4)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", backSpace);
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "a");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", backSpace);
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "b");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-exclamation").length() > 0);
		click(".z-window-highlighted @button:contains(OK)");
		
		click(".z-window.z-window-modal .z-tabpanel:eq(0) .z-combobox-input:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-combobox-open .z-comboitem:eq(4)");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", Keys.BACK_SPACE.toString());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(0)", "2015");
		waitForTime(Setup.getTimeoutL0());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", Keys.BACK_SPACE.toString());
		sendKeys(".z-window.z-window-modal .z-tabpanel:eq(0) .z-textbox:eq(1)", "31");
		waitForTime(Setup.getTimeoutL0());
		click(".z-window.z-window-modal @button:contains('OK')");
		waitForTime(Setup.getTimeoutL0());
		assertTrue(jq(".z-messagebox-exclamation").length() > 0);
		click(".z-window-highlighted @button:contains(OK)");
	}
}





