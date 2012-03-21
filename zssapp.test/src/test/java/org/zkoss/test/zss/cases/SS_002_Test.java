/* SS_002_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 11:04:21 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import java.util.Random;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.Cell;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

import com.google.common.base.Strings;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_002_Test extends ZSSAppTest {

	@Test
	public void open_new_empty_sheet() {
		
		click("$fileMenu");
		click("$newFile:visible");
		
		timeBlocker.waitResponse();
		
    	Random randomGenerator = new Random();
    	Cell cell = getCell(randomGenerator.nextInt(10), randomGenerator.nextInt(10));
    	
    	Assert.assertTrue("cell text shall be empty", Strings.isNullOrEmpty(cell.getText()));
	}
}
