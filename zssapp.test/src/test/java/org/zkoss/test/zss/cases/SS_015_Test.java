/* SS_015_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 14, 2012 5:26:21 PM , Created by sam
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
public class SS_015_Test extends ZSSAppTest {

	@Test
	public void toggle_font_bold() {
		spreadsheet.setSelection(11, 5, 16, 5);
		boolean isFontBold = getCell(11, 5).isFontBold();
		click(".zstbtn-fontBold");
		verifyToggleFontBold(isFontBold, 11, 5, 16, 5);
		
		isFontBold = getCell(11, 5).isFontBold();
		click(".zstbtn-fontBold");
		verifyToggleFontBold(isFontBold, 11, 5, 16, 5);
	}
	
	private void verifyToggleFontBold(boolean b, int tRow, int lCol, int bRow, int rCol) {
		boolean toggle = !b;
		verifyFontBold(toggle, tRow, lCol, bRow, rCol);
	}
}
