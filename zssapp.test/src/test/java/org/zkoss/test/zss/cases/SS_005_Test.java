/* SS_006_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 12:24:41 PM , Created by sam
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
public class SS_005_Test extends ZSSAppTest {

	@Test
	public void open_export_pdf_dialog() {
		Assert.assertFalse("export PDF dialog shall close", isVisible("$_exportToPdfDialog"));
		
    	click("$fileMenu");
    	click("$exportFile");
    	click("$exportPdf");
    	
    	Assert.assertTrue("export PDF dialog shall open", isVisible("$_exportToPdfDialog"));
	}
	
}
