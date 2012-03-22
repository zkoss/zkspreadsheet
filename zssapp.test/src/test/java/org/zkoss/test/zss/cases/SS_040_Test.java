/* SS_040_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 11:33:11 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Assert;
import org.junit.Test;
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
public class SS_040_Test extends ZSSAppTest {
	
	@Test
	public void paste_special_all() {
		int tRow = 15;
		int lCol = 6;
		int bRow = 18;
		int rCol = 7;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		click("$_pasteSpecialDialog $okBtn");
		
		Assert.assertEquals(copyFrom, builder.offset(11, 10).build());
	}
	
	@Test
	public void paste_special_all_except_borders() {
		int tRow = 15;
		int lCol = 6;
		int bRow = 18;
		int rCol = 7;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$allExcpetBorder input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		Assert.assertTrue(copyFrom.equals(builder.offset(11, 10).build(), EqualCondition.EXPECT_BORDER));
	}
	
	@Test
	public void paste_special_column_widths() {
		int tRow = 15;
		int lCol = 6;
		int bRow = 18;
		int rCol = 7;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$colWidth input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		Assert.assertTrue(copyFrom.equals(builder.offset(11, 10).build(), EqualCondition.COLUMN_WIDTH_ONLY));
	}

	@Test
	public void paste_special_formulas() {
		int tRow = 15;
		int lCol = 6;
		int bRow = 18;
		int rCol = 7;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$formula input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		CellCacheAggeration pasteTo = builder.offset(11, 10).build();
		copyFrom.merge(pasteTo, CellCache.Field.VERTICAL_ALIGN, CellCache.Field.HORIZONTAL_ALIGN, CellCache.Field.FONT_COLOR, CellCache.Field.FILL_COLOR, CellCache.Field.BOTTOM_BORDER, CellCache.Field.RIGHT_BORDER);
		Assert.assertTrue(copyFrom.equals(pasteTo, CellCache.EqualCondition.IGNORE_NUMBER_FORMAT));
	}
	
	@Test
	public void paste_special_formulas_and_number_formats() {
		int tRow = 15;
		int lCol = 6;
		int bRow = 18;
		int rCol = 7;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$formulaWithNum input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		CellCacheAggeration pasteTo = builder.offset(11, 10).build();
		copyFrom.merge(pasteTo, CellCache.Field.VERTICAL_ALIGN, CellCache.Field.HORIZONTAL_ALIGN, CellCache.Field.FONT_COLOR, CellCache.Field.FILL_COLOR, CellCache.Field.BOTTOM_BORDER, CellCache.Field.RIGHT_BORDER);
		Assert.assertEquals(copyFrom, pasteTo);	
	}
	
	@Test
	public void paste_special_values() {
		int tRow = 15;
		int lCol = 6;
		int bRow = 18;
		int rCol = 7;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$value input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		CellCacheAggeration pasteTo = builder.offset(11, 10).build();
		copyFrom.merge(pasteTo, CellCache.Field.VERTICAL_ALIGN, CellCache.Field.HORIZONTAL_ALIGN, CellCache.Field.FONT_COLOR, CellCache.Field.FILL_COLOR, CellCache.Field.BOTTOM_BORDER, CellCache.Field.RIGHT_BORDER);
		Assert.assertTrue(copyFrom.equals(pasteTo, EqualCondition.VALUE, EqualCondition.IGNORE_NUMBER_FORMAT));
	}
	
	@Test
	public void paste_special_values_and_number_formats() {
		int tRow = 15;
		int lCol = 6;
		int bRow = 18;
		int rCol = 7;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$valueWithNumFmt input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		CellCacheAggeration pasteTo = builder.offset(11, 10).build();
		copyFrom.merge(pasteTo, CellCache.Field.VERTICAL_ALIGN, CellCache.Field.HORIZONTAL_ALIGN, CellCache.Field.FONT_COLOR, CellCache.Field.FILL_COLOR, CellCache.Field.BOTTOM_BORDER, CellCache.Field.RIGHT_BORDER);
		Assert.assertTrue(copyFrom.equals(pasteTo, EqualCondition.VALUE));
	}

	@Test
	public void paste_special_formats() {
		int tRow = 15;
		int lCol = 6;
		int bRow = 18;
		int rCol = 7;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$fmt input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		CellCacheAggeration pasteTo = builder.offset(11, 10).build();
		copyFrom.merge(pasteTo, CellCache.Field.CELL_TYPE, CellCache.Field.TEXT, CellCache.Field.EDIT);
		Assert.assertEquals(copyFrom, pasteTo);	
	}

	@Test
	public void paste_special_skip_blanks() {
		//K5:K9, 10, 20, <blank>, 40, <blank>
		keyboardDirector.setEditText(5, 10, "10");
		keyboardDirector.setEditText(6, 10, "20");
		keyboardDirector.setEditText(8, 10, "40");
		
		//L5:L9 50, 60, 70, 80, and 90
		keyboardDirector.setEditText(5, 11, 9, 11, new String[]{"50", "60", "70", "80", "90"});
		
		int tRow = 5;
		int lCol = 10;
		int bRow = 9;
		int rCol = 10;
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		
		mouseDirector.openCellContextMenu(5, 11);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$skipBlanks input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		Assert.assertEquals("10", getCell(5, 11).getText());
		Assert.assertEquals("20", getCell(6, 11).getText());
		Assert.assertEquals("70", getCell(7, 11).getText());
		Assert.assertEquals("40", getCell(8, 11).getText());
		Assert.assertEquals("90", getCell(9, 11).getText());
	}

	@Test
	public void paste_special_transpose() {
		CellCache F12 = getCellCache(11, 5);
		CellCache F13 = getCellCache(12, 5);
		spreadsheet.setSelection(11, 5, 12, 5);
		click(".zstbtn-copy");
		
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		jq("$transpose input").getWebElement().click();
		click("$_pasteSpecialDialog $okBtn");
		
		CellCache K12 = getCellCache(11, 10);
		CellCache K13 = getCellCache(11, 11);
		Assert.assertEquals(F12.getText(), K12.getText());
		Assert.assertEquals(F13.getText(), K13.getText());
	}
}
