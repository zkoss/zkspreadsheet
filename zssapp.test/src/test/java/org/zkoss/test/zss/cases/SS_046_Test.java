/* SS_046_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 6:37:25 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.zkoss.test.JavascriptActions;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_046_Test extends ZSSAppTest {

	@Test
	public void formula_editor_select_cell_by_mouse() {
		spreadsheet.focus(14, 10);
		
		WebElement editor = jq(".zsformulabar-editor-real").getWebElement();
		click(".zsformulabar-editor-real");
		editor.sendKeys("=");
		timeBlocker.waitResponse();
		
		spreadsheet.focus(14, 5);
		new JavascriptActions(webDriver)
		.enter(jq(".zsformulabar-editor-real"))
		.perform();

		Assert.assertEquals("125000", getCell(14, 10).getText());
	} 
}
