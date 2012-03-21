/* SS_014_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 14, 2012 5:03:01 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Test;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_014_Test extends ZSSAppTest {

	@Test
	public void set_font_size() {
		spreadsheet.setSelection(11, 5, 16, 5);
		
		click(".zsfontsize .z-combobox-btn");
		click(".zsfontsize-" + 20);
		
		verifyFontSize(20, 11, 5, 16, 5);
	}
}