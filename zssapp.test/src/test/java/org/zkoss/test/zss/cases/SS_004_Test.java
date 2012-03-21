/* SS_004_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 12:18:16 PM , Created by sam
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
public class SS_004_Test extends ZSSAppTest {

	@Test
	public void open_import_file_dialog() {
		Assert.assertFalse("Import dialog shall close", isVisible("$_importFileDialog"));
		
    	click("$fileMenu");
    	click("$importFile");
    	
    	Assert.assertTrue("Import dialog shall open", isVisible("$_importFileDialog"));
	}
}
