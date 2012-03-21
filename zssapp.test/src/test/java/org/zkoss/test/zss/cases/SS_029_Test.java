/* SS_029_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 16, 2012 9:49:26 AM , Created by sam
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
public class SS_029_Test extends ZSSAppTest {

	@Test
	public void toggle_protect() {
		
		keyboardDirector.setEditText(11, 10, "ABC");
		Assert.assertEquals("ABC", getCell(11, 10).getText());
		
		//toggle on
		click(".zschktbtn-protectSheet");
		keyboardDirector.setEditText(11, 10, "DEF");
		//protected: remain ABC
		Assert.assertEquals("ABC", getCell(11, 10).getText());
		
		//toggle off
		click(".zschktbtn-protectSheet");
		keyboardDirector.setEditText(11, 10, "DEF");
		Assert.assertEquals("DEF", getCell(11, 10).getText());
	}
}
