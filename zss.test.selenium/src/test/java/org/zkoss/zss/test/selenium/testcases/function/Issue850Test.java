package org.zkoss.zss.test.selenium.testcases.function;

import static org.junit.Assert.assertTrue;

import org.junit.Test;
import org.openqa.selenium.interactions.Actions;
import org.zkoss.zss.test.selenium.Setup;
import org.zkoss.zss.test.selenium.ZSSTestCase;
import org.zkoss.zss.test.selenium.entity.JQuery;


public class Issue850Test extends ZSSTestCase {
	@Test
	public void testZSS850() throws Exception {
		getTo("/issue3/850-fill-pattern.zul");
		focusSheet().getSheetCtrl();
		
		rightClick(jq(".zsrow:eq(0) .zscell:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-formatCell.z-menuitem"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-window .z-tab:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-window .z-colorbox:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-colorbox-popup:last .z-colorpalette-color:first"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		// press OK
		click(jq(".z-window .z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(
				jq(".zsrow:eq(0) .zscell:eq(1)").css("background-color").equals("rgb(0, 0, 0)"));
		
		for(int i = 1; i < 6; i++) {
			applyFillPattern(
					jq(".zsrow:eq(" + i + ") .zscell:eq(0)"),
					jq(".zsrow:eq(" + i + ") .zscell:eq(1)"), i + 1);
		}
		
		for(int i = 0; i < 6; i++) {
			applyFillPattern(
					jq(".zsrow:eq(" + i + ") .zscell:eq(2)"),
					jq(".zsrow:eq(" + i + ") .zscell:eq(3)"), 7 + i);
		}
		
		for(int i = 0; i < 6; i++) {
			applyFillPattern(
					jq(".zsrow:eq(" + i + ") .zscell:eq(4)"),
					jq(".zsrow:eq(" + i + ") .zscell:eq(5)"), 13 + i);
		}
	}
	
	private void applyFillPattern(JQuery e1, JQuery e2, int nth) {
		rightClick(e2);
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsmenuitem-formatCell.z-menuitem"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-window .z-tab:eq(2)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".z-window .z-combobox-button:eq(1)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".z-combobox-open .z-comboitem:eq(" + nth + ")"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		// press OK
		click(jq(".z-window .z-button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(
			e2.css("background-image").equals(e1.css("background-image")));
	
	}
	
	@Test
	public void testZSS851() throws Exception {
		getTo("/issue3/851-focus-contextmenu.zul");
		focusSheet().getSheetCtrl();
		
		rightClick(jq(".zsrow:eq(0) .zscell:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
				
		click(jq(".zsstylepanel-upper .zstbtn-fontColor"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		new Actions(driver()).sendKeys("test").perform();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsrow:eq(1) .zscell:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(eval("return jq('.zsrow:eq(0) .zscell:eq(0)').zk.$().getText()")
				.equals("test"));
		
		rightClick(jq(".zsrow:eq(1) .zscell:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		click(jq(".zsmenuitem-insertComment"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		sendKeys(".z-window iframe", "test");
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq("@button:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		new Actions(driver()).sendKeys("test").perform();
		waitUntilProcessEnd(Setup.getTimeoutL0());
		click(jq(".zsrow:eq(2) .zscell:eq(0)"));
		waitUntilProcessEnd(Setup.getTimeoutL0());
		
		assertTrue(eval("return jq('.zsrow:eq(1) .zscell:eq(0)').zk.$().getText()")
				.equals("test"));
	}
}





