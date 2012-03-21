/* SS_028_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 2:54:47 PM , Created by sam
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
public class SS_008_Test extends ZSSAppTest {

	@Test
	public void freeze_columns() {
		freezeColumnAndVerify(1);
		freezeColumnAndVerify(2);
		freezeColumnAndVerify(3);
		freezeColumnAndVerify(4);
		
		//toggle_freeze_5_columns()
		
		freezeColumnAndVerify(6);
		freezeColumnAndVerify(7);
		freezeColumnAndVerify(8);
		freezeColumnAndVerify(9);
		freezeColumnAndVerify(10);
	}
	
	@Test
	public void toggle_freeze_5_columns() {
    	click("$viewMenu");
    	click("$freezeCols");
    	click("$freezeCol5");
    	Assert.assertTrue("shall freeze columns", isVisible("div.zsleftblock"));
    	Assert.assertEquals(5, getColumnfreeze());
    	
    	click("$viewMenu");
    	click("$freezeCols");
    	click("$unfreezeCols");
    	Assert.assertFalse("shall unfreeze columns", isVisible("div.zsleftblock"));
    	Assert.assertEquals(0, getColumnfreeze());
	}
	
	private void freezeColumnAndVerify(int columnSize) {
    	click("$viewMenu");
    	click("$freezeCols");
    	click("$freezeCol" + columnSize);
    	Assert.assertTrue("shall freeze columns", isVisible("div.zsleftblock"));
    	Assert.assertEquals(columnSize, getColumnfreeze());
	}
	
}
