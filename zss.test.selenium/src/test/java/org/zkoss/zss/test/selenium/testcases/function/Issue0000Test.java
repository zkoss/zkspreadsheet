package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.CellWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

public class Issue0000Test extends ZSSTestCase {

	@Test
	public void testZSS0005() throws Exception {
		getTo("issue3/5-comment-add-upd-del.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		WebElement b2 = ctrl.getCell("B2").$n();
		JQuery insert = jq(".zsmenuitem-insertComment");
		JQuery edit = jq(".zsmenuitem-editComment");
		JQuery delete = jq(".zsmenuitem-deleteComment");
		
		rightClick(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(insert);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		sendKeys(jq(".z-window:visible iframe"), "test");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-window .z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zscellcomment-popup .z-popup-content").html(), 
			"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:11pt;\">test</span>");
		
		rightClick(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(edit);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		sendKeys(jq(".z-window:visible iframe"), Keys.chord(Keys.CONTROL, "a"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(jq(".z-window:visible iframe"), "test1");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-window .z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zscellcomment-popup .z-popup-content").html(), 
			"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:11pt;\">test1</span>");
		
		rightClick(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(delete);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("none"));
		
		sendKeys(jq("body"), Keys.chord(Keys.CONTROL, "z"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zscellcomment-popup .z-popup-content").html(), 
			"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:11pt;\">test1</span>");
		
		click(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(jq("body"), Keys.chord(Keys.CONTROL, "z"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zscellcomment-popup .z-popup-content").html(), 
			"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:11pt;\">test</span>");
		
		sendKeys(jq("body"), Keys.chord(Keys.CONTROL, "z"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		// 6. protect sheet mode with edit objects allowed can be also worked for above 5 steps, deny otherwise.
		rightClick(jq(".zssheettab:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".zsmenuitem-protectSheet"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-window .z-listitem:last"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-window .z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		// do all things from step 1 ~ 5 again
		rightClick(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(insert);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		sendKeys(jq(".z-window:visible iframe"), "test");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-window .z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zscellcomment-popup .z-popup-content").html(), 
			"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:11pt;\">test</span>");
		
		rightClick(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(edit);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		sendKeys(jq(".z-window:visible iframe"), Keys.chord(Keys.CONTROL, "a"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(jq(".z-window:visible iframe"), "test1");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-window .z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zscellcomment-popup .z-popup-content").html(), 
			"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:11pt;\">test1</span>");
		
		rightClick(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(delete);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertTrue(jq(".zscellcomment-popup").css("display").equals("none"));
		
		sendKeys(jq("body"), Keys.chord(Keys.CONTROL, "z"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zscellcomment-popup .z-popup-content").html(), 
			"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:11pt;\">test1</span>");
		
		click(b2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(jq("body"), Keys.chord(Keys.CONTROL, "z"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		mouseMoveAt(b2, "3,3");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		assertEquals(jq(".zscellcomment-popup .z-popup-content").html(), 
			"<span style=\"font-family:新細明體;color:#000000;font-weight:normal;font-size:11pt;\">test</span>");
		
		sendKeys(jq("body"), Keys.chord(Keys.CONTROL, "z"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
	}
}
