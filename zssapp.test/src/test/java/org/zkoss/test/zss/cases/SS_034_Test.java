/* SS_034_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 16, 2012 11:17:54 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.zkoss.test.JQuery;
import org.zkoss.test.JavascriptActions;
import org.zkoss.test.Keycode;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase 
public class SS_034_Test extends ZSSAppTest {

	@Test
	public void add_sheet() {
		
		int sheetSize = jq(".zssheettab").length();
		click(".zstbtn-addSheet");
		Assert.assertEquals(sheetSize + 1, jq(".zssheettab").length());
	}
	
	@Test
	public void delete_sheet() {
		int sheetSize = jq(".zssheettab").length();
		rightClick(jq(".zssheettab").first());
		click(".zsmenuitem-deleteSheet");
		
		Assert.assertEquals(sheetSize - 1, jq(".zssheettab").length());
	}
	
	@Test
	public void rename_sheet() {
		
		rightClick(jq(".zssheettab").first());
		click(".zsmenuitem-renameSheet");
		
		JQuery $editor = jq(".zssheettab-rename-textbox");
		WebElement editor = $editor.getWebElement();
		editor.clear();
		editor.sendKeys("PP");
		new JavascriptActions(webDriver)
		.keyDown($editor, Keycode.ENTER.intValue())
		.keyUp($editor, Keycode.ENTER.intValue())
		.perform();
		timeBlocker.waitUntil(1);
		
		Assert.assertEquals("PP", jq(".zssheettab-text").first().text());
		
		//rename twice
		rightClick(jq(".zssheettab").first());
		click(".zsmenuitem-renameSheet");
		
		editor = jq(".zssheettab-rename-textbox").getWebElement();
		editor.clear();
		editor.sendKeys("SS");
		new JavascriptActions(webDriver)
		.keyDown($editor, Keycode.ENTER.intValue())
		.keyUp($editor, Keycode.ENTER.intValue())
		.perform();
		timeBlocker.waitUntil(1);
		
		Assert.assertEquals("SS", jq(".zssheettab-text").first().text());
	}
	
	@Test
	public void protect_sheet() {
		
		keyboardDirector.setEditText(11, 10, "ABC");
		timeBlocker.waitUntil(1);
		Assert.assertEquals("ABC", getCell(11, 10).getText());
		
		JQuery target = jq(".zssheettab").first();
		target.getWebElement().click();
		rightClick(target);
		click(".z-menupopup:visible .zsmenuitem-protectSheet");
		
		spreadsheet.focus(11, 10);
		keyboardDirector.setEditText(11, 10, "DEF");
		timeBlocker.waitUntil(1);
		Assert.assertEquals("ABC", getCell(11, 10).getText());
	}
	
	@Test
	public void move_sheet_left() {
		//the first sheet tab cannot allow move left
		rightClick(jq(".zssheettab").first());
		
		String cssName = jq(".zsmenuitem-moveSheetLeft").attr("class");
		Assert.assertTrue(cssName.indexOf("z-menuitem-disd") >= 0);
		
		//move the second sheet left
		JQuery sheet = jq(".zssheettab").eq(1);
		String sheetName = sheet.text();
		click(sheet);
		rightClick(sheet);
		click(".z-menupopup:visible .zsmenuitem-moveSheetLeft");
		
		JQuery firstSheet = jq(".zssheettab").first();
		Assert.assertEquals(sheetName, firstSheet.text());
	}
	
	@Test
	public void move_shift_right() {
		JQuery sheet = jq(".zssheettab").first();
		String sheetName = sheet.text();
		rightClick(sheet);
		click(".z-menupopup:visible .zsmenuitem-moveSheetRight");
		
		Assert.assertEquals(sheetName, jq(".zssheettab").eq(1).text());
	} 
}