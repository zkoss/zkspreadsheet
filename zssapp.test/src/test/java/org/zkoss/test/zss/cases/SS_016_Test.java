/* SS_016_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 14, 2012 5:48:05 PM , Created by sam
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
public class SS_016_Test extends ZSSAppTest{

	@Test
	public void toggle_font_italic() {
		spreadsheet.setSelection(11, 5, 16, 5);
		boolean isFontItalic = getCell(11, 5).isFontItalic();
		click(".zstbtn-fontItalic");
		verifyToggleFontItalic(isFontItalic, 11, 5, 16, 5);
		
		isFontItalic = getCell(11, 5).isFontItalic();
		click(".zstbtn-fontItalic");
		verifyToggleFontItalic(isFontItalic, 11, 5, 16, 5);
	}
	
	private void verifyToggleFontItalic(boolean b, int tRow, int lCol, int bRow, int rCol) {
		boolean toggle = !b;
		verifyFontItalic(toggle, tRow, lCol, bRow, rCol);
	}
}