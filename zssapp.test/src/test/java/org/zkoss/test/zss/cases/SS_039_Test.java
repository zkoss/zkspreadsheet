/* SS_039_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 11:19:53 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Assert;
import org.junit.Test;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.zkoss.test.JavascriptActions;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_039_Test extends ZSSAppTest {

	@Test
	public void upload_file_dialog() {
		
    	click("$fileMenu");
    	click("$openFile");
    	Assert.assertTrue(isVisible("$_openFileDialog"));
    	
	}
	
	@Test
	public void open_file() {
    	click("$fileMenu");
    	click("$openFile");
    	Assert.assertTrue(isVisible("$_openFileDialog"));
    	
    	WebElement e = jq("$_openFileDialog $filesListbox .z-listcell").getWebElement();
    	//TODO: doubleClick on FF 10 doesn't work well 
    	new Actions(webDriver).doubleClick(e).perform();
    	
    	timeBlocker.waitResponse();
		if (browser.isSafari() || browser.isGecko()) {
			timeBlocker.waitUntil(1);
		}
		
		Assert.assertEquals("Collaboration", getCell(3, 0).getText());
	}
}
