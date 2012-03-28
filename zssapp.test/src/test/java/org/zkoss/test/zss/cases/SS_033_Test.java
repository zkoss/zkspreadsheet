/* SS_031_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 16, 2012 10:11:26 AM , Created by sam
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
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_033_Test extends ZSSAppTest {

	@Test
	public void insert_hyperlink() {
		spreadsheet.focus(11, 10);
		
		click(".zstab-insertPanel");
		click(".zstbtn-hyperlink");
		Assert.assertTrue(isVisible("$_insertHyperlinkDialog"));
		
        WebElement inp = jq("$addrCombobox input.z-combobox-inp").getWebElement();
        inp.sendKeys("http://ja.wikipedia.org/wiki");
        if (browser.isIE6() || browser.isIE7()) {
        	timeBlocker.waitUntil(3);
        }
        inp.sendKeys(Keys.ENTER);
        timeBlocker.waitResponse();
        
        JQuery link = getCell(11, 10).jq$n("real").children().first();
        
        Assert.assertEquals("http://ja.wikipedia.org/wiki", link.attr("href"));
	}
	
	@Test
	public void insert_mail_hyperlink() {
		spreadsheet.focus(11, 10);
		
		click(".zstab-insertPanel");
		click(".zstbtn-hyperlink");
		Assert.assertTrue(isVisible("$_insertHyperlinkDialog"));
		click("$mailBtn");
		
        WebElement inp = jq("$mailAddr").getWebElement();
        inp.sendKeys("example@potix.com");
        timeBlocker.waitUntil(1);
        inp.sendKeys(Keys.TAB);
        timeBlocker.waitResponse();
        click("$_insertHyperlinkDialog $okBtn");
        
        JQuery link = getCell(11, 10).jq$n("real").children().first();
        Assert.assertEquals("mailto:example@potix.com", link.attr("href"));
	}
	
	@Test
	public void insert_doc_hyperlink() {
		spreadsheet.focus(11, 10);
		
		click(".zstab-insertPanel");
		click(".zstbtn-hyperlink");
		Assert.assertTrue(isVisible("$_insertHyperlinkDialog"));
		click("$docBtn");
		
		click("$refSheet .z-treecell:eq(0)");
		click("$okBtn");
		
		JQuery link = getCell(11, 10).jq$n("real").children().first();
		Assert.assertTrue(link.attr("href").lastIndexOf("Input!A1") >= 0);
	}
}
