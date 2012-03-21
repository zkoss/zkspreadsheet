/* SS_003_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 12:12:20 PM , Created by sam
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
public class SS_003_Test extends ZSSAppTest {
	
	@Test
	public void open_file_dialog() {
		Assert.assertFalse("File dialog shall close", isVisible("$_openFileDialog"));
		
    	click("$fileMenu");
    	click("$openFile");
    	
    	Assert.assertTrue("File dialog shall open", isVisible("$_openFileDialog"));
	} 
}
