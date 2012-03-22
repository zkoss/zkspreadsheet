/* SS_045_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 3:21:41 PM , Created by sam
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
public class SS_010_Test extends ZSSAppTest {

	@Test
	public void close_book() {
		Assert.assertTrue(hasBook());
		
		click(".zstbtn-closeBook");
		
		Assert.assertFalse(hasBook());
	}
}
