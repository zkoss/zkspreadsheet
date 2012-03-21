/* SS_017_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 14, 2012 5:55:03 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import java.util.Iterator;

import org.junit.Assert;
import org.junit.Test;
import org.zkoss.test.zss.Cell;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_017_Test extends ZSSAppTest {

	@Test
	public void toggole_font_underline() {
		spreadsheet.setSelection(11, 5, 16, 5);
		boolean isFontUnderline = getCell(11, 5).isFontUnderline();
		click(".zstbtn-fontUnderline");
		verifyToggleFontUnderline(isFontUnderline, 11, 5, 16, 5);
		
		isFontUnderline = getCell(11, 5).isFontUnderline();
		click(".zstbtn-fontUnderline");
		verifyToggleFontUnderline(isFontUnderline, 11, 5, 16, 5);
	}
	
	private void verifyToggleFontUnderline(boolean b, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> cs = iterator(tRow, lCol, bRow, rCol);
		boolean toggle = !b;
		while (cs.hasNext()) {
			Cell c = cs.next();
			Assert.assertEquals(toggle, c.isFontUnderline());
		}
	}
}
