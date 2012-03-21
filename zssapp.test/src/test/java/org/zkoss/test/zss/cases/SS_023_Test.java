/* SS_023_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 15, 2012 11:39:54 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import java.util.Iterator;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.Cell;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_023_Test extends ZSSAppTest {
	
	//TODO: wraped cell shall also check row height
	@Test
	public void toggle_wrap() {
		spreadsheet.setSelection(20, 1, 23, 1);
		
		boolean wrap = getCell(20, 1).isWrap();
		click(".zstbtn-wrapText");
		verifyWrap(wrap, 20, 1, 23, 1);
		
		wrap = getCell(20, 1).isWrap();
		click(".zstbtn-wrapText");
		verifyWrap(wrap, 20, 1, 23, 1);
	}
	
	private void verifyWrap(boolean b, int tRow, int lCol, int bRow, int rCol) {
		boolean toggle = !b;
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertEquals(toggle, c.isWrap());
		}
	}
}
