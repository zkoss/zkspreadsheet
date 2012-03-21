/* SS_025_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 12:28:41 PM , Created by sam
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
public class SS_006_Test extends ZSSAppTest {

	@Test
	public void toggle_formula_bar() {
		
    	click("$viewMenu");
    	click("$viewFormulaBar");
    	Assert.assertFalse("Formula bar shall be invisible", isVisible(".zsformulabar"));
    	
    	click("$viewMenu");
    	click("$viewFormulaBar a.z-menu-item-cnt-unck");
    	Assert.assertTrue("Formula bar shall be visible", isVisible(".zsformulabar"));
	}
}
