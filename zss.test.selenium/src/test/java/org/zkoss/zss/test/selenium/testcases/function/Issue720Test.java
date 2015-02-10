package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue720Test extends ZSSTestCase {

	@Test
	public void testZSS725() throws Exception {
		getTo("issue3/725-richtext-editbox.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		for(int i = 0; i < 9; i++) {
			rightClick(ctrl.getCell(i, 0));
			waitUntilProcessEnd(Setup.getTimeoutL0());
			click(jq(".zsmenuitem-richTextEdit.z-menuitem"));
			waitUntilProcessEnd(Setup.getTimeoutL0());
			
			JQuery iframe = jq(".z-window iframe"); 
			sendKeys(iframe, "abc");
			sendKeys(iframe, Keys.chord(Keys.CONTROL, "a"));
			waitUntilProcessEnd(Setup.getTimeoutL0());
			click(jq(".cleditorButton:eq(" + i + ")"));
			waitUntilProcessEnd(Setup.getTimeoutL0());
			
			if(i == 6) {
				click(jq(".cleditorPopup.cleditorList>div:visible:eq(0)"));
				waitUntilProcessEnd(Setup.getTimeoutL0());
			} else if(i == 7) {
				click(jq(".cleditorPopup.cleditorList>div:visible:eq(15)"));
				waitUntilProcessEnd(Setup.getTimeoutL0());
			} else if(i == 8) {
				click(jq(".z-colorpalette:visible span:eq(5)"));
				waitUntilProcessEnd(Setup.getTimeoutL0());
			}
			
			click(jq(".z-window .z-button:eq(0)"));
			waitUntilProcessEnd(Setup.getTimeoutL0());
		}
		
		assertEquals(jq(".zsrow:eq(0) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:新細明體;color:#000000;font-weight:bold;font-size:11pt;\">abc</span>");
		
		assertEquals(jq(".zsrow:eq(1) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-style:italic;font-size:11pt;\">abc</span>");	
		
		assertEquals(jq(".zsrow:eq(2) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:新細明體;color:#000000;text-decoration: underline;font-weight:normal;font-size:11pt;\">abc</span>");
		
		assertEquals(jq(".zsrow:eq(3) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:新細明體;color:#000000;text-decoration: line-through;font-weight:normal;font-size:11pt;\">abc</span>");	
		
		assertEquals(jq(".zsrow:eq(4) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:8pt;vertical-align:sub;\">abc</span>");
		
		assertEquals(jq(".zsrow:eq(5) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:8pt;vertical-align:super;\">abc</span>");	
		
		assertEquals(jq(".zsrow:eq(6) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:arial;color:#000000;font-weight:normal;font-size:11pt;\">abc</span>");
		
		assertEquals(jq(".zsrow:eq(7) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:72pt;\">abc</span>");	
		
		assertEquals(jq(".zsrow:eq(8) .zscell:eq(0) .zscelltxt-real").html().trim(), 
				"<span style=\"font-family:新細明體;color:#0000ff;font-weight:normal;font-size:11pt;\">abc</span>");	
		
				
	}
	
	@Ignore("Need export")
	@Test
	public void testZSS728() throws Exception {
		
	}
	
	@Test
	public void testZSS729() throws Exception {
		getTo("/issue3/729-validation.zul");
		
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		SheetCtrlWidget sheetCtrl = spreadsheet.getSheetCtrl();
		click(jq("@button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		CellWidget A1 = sheetCtrl.getCell(0, 0);
		click(A1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(ZSStyle.DROPDOWN_BTN));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("abc", jq(".zsdv-popup-cave").text());
		click(jq("@button:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		CellWidget B1 = sheetCtrl.getCell(0, 1);
		JQuery dropdown = jq(ZSStyle.DROPDOWN_BTN);
		click(B1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(dropdown);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals("def", jq(".zsdv-popup-cave").text());
		click(jq("@button:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		spreadsheet.focus();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(A1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse(dropdown.exists());
		click(B1);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertFalse(dropdown.exists());
	}
}
