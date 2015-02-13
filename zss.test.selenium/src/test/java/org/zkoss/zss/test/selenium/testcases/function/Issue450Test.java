package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertEquals;

import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.test.selenium.AssertUtil;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.SheetCtrlWidget;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class Issue450Test extends ZSSTestCase {
	@Test
	public void testZSS450() throws Exception {
		getTo("/issue3/450-delete-row.zul");

		click(jq("@button:eq(3)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(0)"));
		waitForTime(Setup.getTimeoutL0());
		click(jq("@button:eq(1)"));
		waitForTime(Setup.getTimeoutL1());
		click(jq("@button:eq(2)"));
		waitForTime(Setup.getTimeoutL0());
		assertEquals("A140", getSpreadsheet().getSheetCtrl().getCell("A140").getText());
		assertEquals("D140", getSpreadsheet().getSheetCtrl().getCell("D140").getText());
	}
	
	@Test
	public void testZSS451() throws Exception {
		getTo("issue3/451-delete-column.zul");
		
		String prev = "100";
		for(int i = 0; i < 100; i++) {
			eval("jq('.zsscroll').scrollLeft(" + (i * 100) + ")");
			String current = eval("return jq('.zsscroll').scrollLeft()").toString();
			if(prev.equals(current))
				break;
			else
				prev = current;
		}
		
		prev = "100";
		for(int i = 0; i < 100; i++) {
			eval("jq('.zsscroll').scrollTop(" + (i * 100) + ")");
			String current = eval("return jq('.zsscroll').scrollTop()").toString();
			if(prev.equals(current))
				break;
			else
				prev = current;
		}
		
		waitForTime(Setup.getTimeoutL0());
		click(".z-button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		
		for(int i = Integer.parseInt(prev); i >= -100; i -= 100) {
			eval("jq('.zsscroll').scrollTop(" + i + ")");
		}
		
		waitForTime(Setup.getTimeoutL1());
		AssertUtil.assertNoJSError();
	}
	
	@Test
	public void testZSS452() throws Exception {
		getTo("issue3/452-insertAtAnotherSheet.zul");
		SpreadsheetWidget ss = focusSheet();
		SheetCtrlWidget ctrl = ss.getSheetCtrl();
		
		click(".z-button:eq(0)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-button:eq(1)");
		waitForTime(Setup.getTimeoutL0());
		click(".z-button:eq(2)");
		waitForTime(Setup.getTimeoutL0());
		
		assertEquals(ctrl.getCell("A2").getText(), "");
	}
	
	@Ignore("Hyperlink")
	@Test
	public void testZSS454() throws Exception {
	}
	
	@Ignore("Hyperlink")
	@Test
	public void testZSS457() throws Exception {
	}
}
