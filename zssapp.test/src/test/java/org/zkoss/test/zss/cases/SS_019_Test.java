/* SS_019_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 14, 2012 6:19:48 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Test;
import org.zkoss.test.Color;
import org.zkoss.test.JQuery;
import org.zkoss.test.zss.CellCacheAggeration;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_019_Test extends ZSSAppTest {

	
	@Test
	public void border_by_toolbarbutton() {
		//default border color
		int tRow = 11;
		int lCol = 10;
		int bRow = 16;
		int rCol = 11;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		CellCacheAggeration borderComparsionCache = builder.expand(1).build();
		
		click(".zstbtn-border");
		
		verifyBorder(BorderType.BOTTOM, "#000000", tRow, lCol, bRow, rCol, borderComparsionCache);
	}
	
	@Test
	public void border_bottom() {
		setBorderWithColorAndVerify(BorderType.BOTTOM, 11, 10, 16, 11);
	}
	
	@Test
	public void border_top() {
		setBorderWithColorAndVerify(BorderType.TOP, 11, 10, 16, 11);
	}
	
	@Test
	public void border_left() {
		setBorderWithColorAndVerify(BorderType.LEFT, 11, 10, 16, 11);
	}
	
	@Test
	public void border_right() {
		setBorderWithColorAndVerify(BorderType.RIGHT, 11, 10, 16, 11);
	}
	
	@Test
	public void border_no() {
		setBorderWithColorAndVerify(BorderType.NO, 11, 8, 16, 9);
	}
	
	@Test
	public void border_all() {
		setBorderWithColorAndVerify(BorderType.ALL, 11, 8, 16, 9);
	}
	
	@Test
	public void border_outside() {
		setBorderWithColorAndVerify(BorderType.OUTSIDE, 11, 8, 16, 9);
	}
	
	@Test
	public void border_inside() {
		setBorderWithColorAndVerify(BorderType.INSIDE, 11, 8, 16, 9);
	}
	
	@Test
	public void border_inside_horizontal() {
		setBorderWithColorAndVerify(BorderType.INSIDE_HORIZONTAL, 11, 8, 16, 9);
	}
	
	@Test
	public void border_inside_vertical() {
		setBorderWithColorAndVerify(BorderType.INSIDE_VERTICAL, 11, 8, 16, 9);
	}
	
	private void setBorderWithColorAndVerify(BorderType border, int tRow, int lCol, int bRow, int rCol) {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		CellCacheAggeration borderComparsionCache = builder.expand(1).build();
		Color color = selectAnotherBorderColor();
		
		click(".zstbtn-border .zstbtn-arrow");
		click(".zsmenuitem-" + border.toString());
		
		verifyBorder(border, color.getColor(), tRow, lCol, bRow, rCol, borderComparsionCache);
	}
	
	Color selectAnotherBorderColor() {
		click(".zstbtn-border .zstbtn-arrow");
		click(".zscolormenu");
		
		JQuery target = jqFactory.create("'.z-colorpalette-colorbox'").eq(68);
		String color = target.css("background-color");
		click(target);
		
		return new Color(target.css("background-color"));
	}
}
