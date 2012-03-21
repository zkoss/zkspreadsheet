/* SS_047_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 4:05:50 PM , Created by sam
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
public class SS_011_Test extends ZSSAppTest {

	@Test
	public void click_export_PDF_button() {
		Assert.assertFalse("export PDF dialog shall close", isVisible("$_exportToPdfDialog"));
		
		click(".zstbtn-exportPDF");
		
		Assert.assertTrue("shall open export PDF dialog", isVisible("$_exportToPdfDialog"));
	}
}
