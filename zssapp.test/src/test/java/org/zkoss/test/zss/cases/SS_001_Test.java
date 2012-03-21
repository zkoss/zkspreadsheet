/* SS_001_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 8:40:39 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_001_Test extends ZSSAppTest {

	@Test
	public void click_file_menu() {
		Assert.assertFalse(spreadsheet.isSelection(5, 5));
		
		spreadsheet.focus(5, 5);
		click("$fileMenu");
		Assert.assertTrue("Focus shall remain at same cell", spreadsheet.isSelection(5, 5));
	}
	
	@Test
	public void click_view_menu() {
		Assert.assertFalse(spreadsheet.isSelection(5, 5));
		
		spreadsheet.focus(5, 5);
		click("$viewMenu");
		
		Assert.assertTrue("Focus shall remain at same cell", spreadsheet.isSelection(5, 5));
	}
}
