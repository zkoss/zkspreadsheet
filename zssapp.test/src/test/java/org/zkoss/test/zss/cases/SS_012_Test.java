/* SS_050_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 4:10:31 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.Cell;
import org.zkoss.test.zss.CellCache;
import org.zkoss.test.zss.CellCache.EqualCondition;
import org.zkoss.test.zss.CellCacheAggeration;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_012_Test extends ZSSAppTest {

	@Test
	public void paste_all_by_button() {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(11, 5, 16, 5);
		CellCacheAggeration copyFrom = builder.build();
		prepareCopySource(copyFrom.getTop(), copyFrom.getLeft(), copyFrom.getBottom(), copyFrom.getRight());
		
		focus(11, 10);
		click(".zstbtn-paste");
		
		CellCacheAggeration pasteTo = builder.offset(11, 10).build();
		Assert.assertEquals(copyFrom, pasteTo);
	}
	
	@Test
	public void paste_all_by_menuitem() {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(11, 5, 16, 5);
		CellCacheAggeration copyFrom = builder.build();
		prepareCopySource(copyFrom.getTop(), copyFrom.getLeft(), copyFrom.getBottom(), copyFrom.getRight());
		
		focus(11, 10);
		click(".zstbtn-paste .zstbtn-arrow");
		click(".zsmenuitem-paste");
		
		//test paste twice
		focus(20, 10);
		click(".zstbtn-paste .zstbtn-arrow");
		click(".zsmenuitem-paste");
		
		CellCacheAggeration pasteTo = builder.offset(20, 10).build();
		Assert.assertEquals(copyFrom, pasteTo);
	}
	
	@Test
	public void paste_formula() {
		pasteFormulaAndVerify(11, 5, 16, 5, 16, 6);
	}
	
	@Test
	public void paste_value() {
		pasteValueAndVerify(11, 5, 16, 5, 20, 8);
	}
	
	@Test
	public void paste_all_expect_border() {
		//TODO: enhance border comparison
		pasteAllExpectBorderAndVerify(11, 5, 16, 5, 11, 10); //paste at empty cells
//		pasteAllExpectBorderAndVerify(11, 5, 16, 5, 16, 6); //paste at cells that contains value
	}
	
	@Test
	public void paste_transpose() {
		CellCache F12 = getCellCache(11, 5);
		CellCache F13 = getCellCache(12, 5);
		spreadsheet.setSelection(11, 5, 12, 5);
		click(".zstbtn-copy");
		
		focus(11, 10);
		click(".zstbtn-paste .zstbtn-arrow");
		click(".zsmenuitem-pasteTranspose");
		
		CellCache K12 = getCellCache(11, 10);
		CellCache K13 = getCellCache(11, 11);
		Assert.assertEquals(F12.getText(), K12.getText());
		Assert.assertEquals(F13.getText(), K13.getText());
	}

	@Test
	public void open_paste_special_dialog() {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(11, 5, 12, 5);
		CellCacheAggeration copyFrom = builder.build();
		spreadsheet.setSelection(11, 5, 12, 5);
		click(".zstbtn-copy");
		click(".zstbtn-paste .zstbtn-arrow");
		
		focus(11, 10);
		click(".zsmenuitem-pasteSpecial");
		
		Assert.assertTrue(isVisible("$_pasteSpecialDialog"));
		click("$_pasteSpecialDialog $okBtn");
		CellCacheAggeration pasteTo = builder.offset(11, 10).build();
		
		Assert.assertEquals(copyFrom, pasteTo);
	}
	
	@Test
	public void paste_from_cut() {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(11, 5, 12, 5);
		CellCacheAggeration copyFrom = builder.build();
		Assert.assertFalse(spreadsheet.isHighlight(11, 5, 12, 5));
		
		spreadsheet.setSelection(11, 5, 12, 5);
		click(".zstbtn-cut");
		
		Assert.assertTrue(spreadsheet.isHighlight(11, 5, 12, 5));
		
		focus(11, 10);
		click(".zstbtn-paste");
		
		CellCacheAggeration pasteTo = builder.offset(11, 10).build();
		Assert.assertEquals(copyFrom, pasteTo);
		
		//clear copy from
		for (CellCache c : copyFrom) {
			Assert.assertEquals(Cell.CellType.BLANK, getCellType(c.getRow(), c.getCol()));
		}
	}

	private void prepareCopySource(int tRow, int lCol, int bRow, int rCol) {
		Assert.assertFalse(spreadsheet.isHighlight(tRow, lCol, bRow, rCol));
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-copy");
		
		Assert.assertTrue(spreadsheet.isHighlight(tRow, lCol, bRow, rCol));
	}
	
	private void pasteAllExpectBorderAndVerify(int tRow, int lCol, int bRow, int rCol, int pasteToRow, int pasteToCol) {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
//		CellCacheAggeration pasteDestination = builder.offset(pasteToRow, pasteToCol).build();
//		copyFrom.merge(pasteDestination, CellCache.Field.BOTTOM_BORDER, CellCache.Field.RIGHT_BORDER);
		
		prepareCopySource(copyFrom.getTop(), copyFrom.getLeft(), copyFrom.getBottom(), copyFrom.getRight());
		
		focus(pasteToRow, pasteToCol);
		click(".zstbtn-paste .zstbtn-arrow");
		click(".zsmenuitem-pasteAllExceptBorder");
		
		CellCacheAggeration pasteTo = builder.offset(pasteToRow, pasteToCol).build();
		Assert.assertTrue(copyFrom.equals(pasteTo, EqualCondition.EXPECT_BORDER));
	}
	
	//paste formula remain original's style
	private void pasteFormulaAndVerify(int tRow, int lCol, int bRow, int rCol, int pasteToRow, int pasteToCol) {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		CellCacheAggeration pasteDestination = builder.offset(pasteToRow, pasteToCol).build();
		//when paste formula, it remain same style: align color, border etc
		copyFrom.merge(pasteDestination, CellCache.Field.VERTICAL_ALIGN, CellCache.Field.HORIZONTAL_ALIGN, CellCache.Field.FONT_COLOR, CellCache.Field.FILL_COLOR, CellCache.Field.BOTTOM_BORDER, CellCache.Field.RIGHT_BORDER);
		
		prepareCopySource(copyFrom.getTop(), copyFrom.getLeft(), copyFrom.getBottom(), copyFrom.getRight());
		
		focus(pasteToRow, pasteToCol);
		click(".zstbtn-paste .zstbtn-arrow");
		click(".zsmenuitem-pasteFormula");
		
		CellCacheAggeration pasteTo = builder.offset(pasteToRow, pasteToCol).build();
		Assert.assertEquals(copyFrom, pasteTo);
	}
	
	private void pasteValueAndVerify(int tRow, int lCol, int bRow, int rCol, int pasteToRow, int pasteToCol) {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		CellCacheAggeration pasteDestination = builder.offset(pasteToRow, pasteToCol).build();
		//remain style
		copyFrom.merge(pasteDestination, CellCache.Field.VERTICAL_ALIGN, CellCache.Field.HORIZONTAL_ALIGN, CellCache.Field.FONT_COLOR, CellCache.Field.FILL_COLOR, CellCache.Field.BOTTOM_BORDER, CellCache.Field.RIGHT_BORDER);
		
		prepareCopySource(copyFrom.getTop(), copyFrom.getLeft(), copyFrom.getBottom(), copyFrom.getRight());
		
		focus(pasteToRow, pasteToCol);
		click(".zstbtn-paste .zstbtn-arrow");
		click(".zsmenuitem-pasteValue");
		
		CellCacheAggeration pasteTo = builder.offset(pasteToRow, pasteToCol).build();
		Assert.assertTrue(copyFrom.equals(pasteTo, EqualCondition.VALUE));
	}
}
