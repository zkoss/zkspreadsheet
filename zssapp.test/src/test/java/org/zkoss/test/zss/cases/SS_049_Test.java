/* SS_049_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 2, 2012 9:08:27 AM , Created by sam
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
public class SS_049_Test extends ZSSAppTest {

	@Test
	public void formula_SLN() {
		keyboardDirector.setEditText(11, 10, "=SLN(10000, 5000, 5)");
		
		Assert.assertEquals("1000", getCell(11, 10).getText());
	}
}
