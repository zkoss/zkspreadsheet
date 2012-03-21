/* SS_030_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 16, 2012 10:00:17 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Assert;
import org.junit.Test;
import org.zkoss.test.Border;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_030_Test extends ZSSAppTest {

	@Test
	public void toggle_gridline() {
		
		//toggle on
		click(".zschktbtn-gridlines");
		Border expected = new Border("1px", "solid", "#D0D7E9");
		Assert.assertEquals(expected, getCell(11, 11).getRightBorder());
		
		//toggle off
		click(".zschktbtn-gridlines");
		expected = new Border("1px", "solid", "#FFFFFF");
		Assert.assertEquals(expected, getCell(11, 11).getRightBorder());
	}
}
