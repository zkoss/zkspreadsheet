/* SS_022_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 15, 2012 11:01:31 AM , Created by sam
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
public class SS_022_Test extends ZSSAppTest {
	
	//TODO: also check overflow
	@Test
	public void vertical_align_top () {
		setVerticalAlignAndVerify(AlignType.VERTICAL_ALIGN_TOP, 11, 5, 16, 6);
	}
	
	@Test
	public void vertical_align_middle() {
		setVerticalAlignAndVerify(AlignType.VERTICAL_ALIGN_MIDDLE, 11, 5, 16, 6);
	}
	
	@Test
	public void vertical_align_bottom() {
		setVerticalAlignAndVerify(AlignType.VERTICAL_ALIGN_BOTTOM, 11, 5, 16, 6);
	}
	
	@Test
	public void horizontal_align_left() {
		setHorizontalAlignAndVerify(AlignType.HORIZONTAL_ALIGN_LEFT, 11, 5, 16, 6);
	}
	
	@Test
	public void horizontal_align_center() {
		setHorizontalAlignAndVerify(AlignType.HORIZONTAL_ALIGN_CENTER, 11, 5, 16, 6);
	}
	
	@Test
	public void horizontal_align_right() {
		setHorizontalAlignAndVerify(AlignType.HORIZONTAL_ALIGN_RIGHT, 11, 5, 16, 6);
	}
	
	private void setVerticalAlignAndVerify(AlignType align, int tRow, int lCol, int bRow, int rCol) {
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-verticalAlign .zstbtn-arrow");
		click(".zsmenuitem-" + align.toString());
		
		verifyAlign(align, tRow, lCol, bRow, rCol);
	}
	
	private void setHorizontalAlignAndVerify(AlignType align, int tRow, int lCol, int bRow, int rCol) {
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-horizontalAlign .zstbtn-arrow");
		click(".zsmenuitem-" + align.toString());
		
		verifyAlign(align, tRow, lCol, bRow, rCol);
	}
}
