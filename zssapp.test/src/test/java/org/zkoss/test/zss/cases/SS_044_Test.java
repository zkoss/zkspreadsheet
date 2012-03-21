/* SS_044_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 5:12:26 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.Rect;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_044_Test extends ZSSAppTest {
	
	@Test
	public void select_by_shift() {
		Assert.assertFalse(isVisible(".zsselect"));
		
		keyboardDirector.shiftSelect(11, 5, 15, 8);
		
		Assert.assertTrue(isVisible(".zsselect"));
		Rect sel = spreadsheet.getSheetCtrl().getLastSelection();
		Assert.assertEquals(11, sel.getTop());
		Assert.assertEquals(5, sel.getLeft());
		Assert.assertEquals(15, sel.getBottom());
		Assert.assertEquals(8, sel.getRight());
	}
}
