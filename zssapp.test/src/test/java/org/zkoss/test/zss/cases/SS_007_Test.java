/* SS_027_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 12:40:09 PM , Created by sam
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
public class SS_007_Test extends ZSSAppTest {

	@Test
	public void freeze_rows() {
		freezeRowAndVerify(1);
		freezeRowAndVerify(2);
		freezeRowAndVerify(3);
		freezeRowAndVerify(4);
		
		//toggle_freeze_5_rows()
		
		freezeRowAndVerify(6);
		freezeRowAndVerify(7);
		freezeRowAndVerify(8);
		freezeRowAndVerify(9);
		freezeRowAndVerify(10);
	}
	
	@Test
	public void toggle_freeze_5_rows() {
    	click("$viewMenu");
    	click("$freezeRows");
    	click("$freezeRow5");
    	Assert.assertTrue("shall freeze rows", isVisible("div.zstopblock"));
    	Assert.assertEquals(5, getRowfreeze());
    	
    	click("$viewMenu");
    	click("$freezeRows");
    	click("$unfreezeRows");
    	Assert.assertFalse("shall unfreeze rows", isVisible("div.zstopblock"));
    	Assert.assertEquals(0, getRowfreeze());
	}
	
	private void freezeRowAndVerify(int rowSize) {
    	click("$viewMenu");
    	click("$freezeRows");
    	click("$freezeRow" + rowSize);
    	Assert.assertTrue("shall freeze rows", isVisible("div.zstopblock"));
    	Assert.assertEquals(rowSize, getRowfreeze());
	}
}
