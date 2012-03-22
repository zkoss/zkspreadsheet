/* SS_032_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 16, 2012 11:06:34 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Assert;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.zkoss.test.JQuery;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_032_Test extends ZSSAppTest {

	@Test
	public void open_insert_formula_dialog() {
		
		click(".zsformulabar-formulabtn");
		Assert.assertTrue(isVisible("$_insertFormulaDialog"));
	} 
	
	@Test
	public void formulabar_editor() {
		
		spreadsheet.focus(12, 10);
		JQuery formulaBar = jq(".zsformulabar-editor-real");
		formulaBar.getWebElement().click();
		formulaBar.getWebElement().sendKeys("123456");
		formulaBar.getWebElement().sendKeys(Keys.ENTER);
		timeBlocker.waitUntil(1);
		
		Assert.assertEquals("123456", getCell(12, 10).getText());
	}
}
