/* SS_036_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 16, 2012 4:47:16 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Assert;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.zkoss.test.Color;
import org.zkoss.test.JQuery;
import org.zkoss.test.zss.CellCacheAggeration;
import org.zkoss.test.zss.Header;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase 
public class SS_036_Test extends ZSSAppTest {
	
	@Test
	public void row_font_family() {
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zsfontfamily .z-combobox-btn");
		click(".zsfontfamily-arial");
		
		verifyFontFamily("arial", 12, 5, 12, 9);
	}
	
	@Test
	public void row_font_size() {
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zsfontsize .z-combobox-btn");
		click(".zsfontsize-20");
		
		verifyFontSize(20, 12, 5, 12, 9);
	}
	
	@Test
	public void row_font_bold() {
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zstbtn-fontBold");
		
		verifyFontBold(true, 12, 5, 12, 9);	
	}
	
	@Test
	public void row_font_italic() {
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zstbtn-fontItalic");
		
		verifyFontItalic(true, 12, 5, 12, 9);
	}
	
	@Test
	public void row_font_color() {
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zstbtn-fontColor .zstbtn-cave");
		JQuery target = jqFactory.create("'.z-colorpalette-colorbox'").eq(68);
		String color = target.css("background-color");
		click(target);
		
		verifyFontColor(new Color(color), 12, 5, 12, 9);
	}
	
	@Test
	public void row_fill_color() {
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zstbtn-fillColor .zstbtn-cave");
		JQuery target = jqFactory.create("'.z-colorpalette-colorbox'").eq(68);
		String color = target.css("background-color");
		click(target);
		
		verifyFillColor(new Color(color), 12, 5, 12, 9);
	}
	
	@Test
	public void row_border() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 12;
		int rCol = 9;
		CellCacheAggeration cache = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zstbtn-border .zstbtn-arrow");
		click(".zsmenuitem-" + BorderType.BOTTOM.toString());
		
		verifyBorder(BorderType.BOTTOM, "#000000", tRow, lCol, bRow, rCol, cache);
	}
	
	@Test
	public void row_vertical_align() {
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zstbtn-verticalAlign .zstbtn-arrow");
		click(".zsmenuitem-" + AlignType.VERTICAL_ALIGN_TOP.toString());
		
		verifyAlign(AlignType.VERTICAL_ALIGN_TOP, 12, 5, 12, 9);
	}
	
	@Test
	public void row_horizontal_align() {
		mouseDirector.openRowContextMenu(12);
		
		click(".zsstylepanel .zstbtn-horizontalAlign .zstbtn-arrow");
		click(".zsmenuitem-" + AlignType.HORIZONTAL_ALIGN_LEFT.toString());
		
		verifyAlign(AlignType.HORIZONTAL_ALIGN_LEFT, 12, 5, 12, 9);
	}
	
	@Test
	public void row_cut() {
		CellCacheAggeration from = getCellCacheAggerationBuilder(12, 5, 12, 9).build();
		
		mouseDirector.openRowContextMenu(12);
		click(".z-menu-popup:visible .zsmenuitem-cut");
		
		mouseDirector.openRowContextMenu(16);
		click(".z-menu-popup:visible .zsmenuitem-paste");
		
		verifyPasteAll(PasteSource.CUT, from, 16, 5, 16, 9);
	}
	
	@Test
	public void row_copy() {
		CellCacheAggeration from = getCellCacheAggerationBuilder(12, 5, 12, 9).build();
		
		mouseDirector.openRowContextMenu(12);
		click(".z-menu-popup:visible .zsmenuitem-copy");
		
		mouseDirector.openRowContextMenu(16);
		click(".z-menu-popup:visible .zsmenuitem-paste");
		
		verifyPasteAll(PasteSource.COPY, from, 16, 5, 16, 9);
	}
	
	@Test
	public void row_paste_special() {
		CellCacheAggeration from = getCellCacheAggerationBuilder(12, 5, 12, 9).build();
		mouseDirector.openRowContextMenu(12);
		click(".z-menu-popup:visible .zsmenuitem-copy");
		
		mouseDirector.openRowContextMenu(16);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		Assert.assertTrue(isVisible("$_pasteSpecialDialog"));
		click("$_pasteSpecialDialog $okBtn");
		
		verifyPasteAll(PasteSource.COPY, from, 12, 5, 12, 9);
	}
	
	@Test
	public void row_insert() {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(13, 5, 13, 9);
		CellCacheAggeration from = builder.build();
		mouseDirector.openRowContextMenu(13);
		
		click(".z-menu-popup:visible .zsmenuitem-insertSheetRow");
		
		verifyInsert(Insert.CELL_DOWN, from, builder);
	}

	@Test
	public void row_delete() {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(13, 5, 13, 9);
		CellCacheAggeration from = builder.down().build();
		mouseDirector.openRowContextMenu(13);
		
		click(".z-menu-popup:visible .zsmenuitem-deleteSheetRow");
		
		verifyDelete(Delete.CELL_UP, from, builder, null);
	}

	
	@Test
	public void row_clear_content() {
		CellCacheAggeration cache = getCellCacheAggerationBuilder(12, 5, 12, 9).build();
		mouseDirector.openRowContextMenu(12);
		click(".z-menu-popup:visible .zsmenuitem-clearContent");
		verifyClearContent(cache);
	}
	
	@Test
	public void row_format_number() {
		
		mouseDirector.openRowContextMenu(12);
		click(".z-menu-popup:visible .zsmenuitem-formatCell");
		Assert.assertTrue(isVisible("$_formatNumberDialog"));
		
		click("@listcell[label=\"Accounting\"] div.z-overflow-hidden");
		click("@listcell[label=\"$1,234\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		verifyFormatNumber(FormatNumber.ACCOUNTING, 12, 5, 12, 9);
	}
	
	@Test
	public void row_height() {
		mouseDirector.openRowContextMenu(12);
		
		click(".z-menu-popup:visible .zsmenuitem-rowHeight");
		WebElement input =  jq("$headerSize").getWebElement();
		input.clear();
		input.sendKeys("40");
		input.sendKeys(Keys.ENTER);
		timeBlocker.waitUntil(1);
		timeBlocker.waitResponse();
		
		int rowHeight = getCell(12, 5).jq$n().height();
		Assert.assertTrue(rowHeight >= 40 - 2);
		Assert.assertTrue(rowHeight <= 40 + 2);
	}
	
	@Test
	public void toggle_row_hide() {
		mouseDirector.openRowContextMenu(12);
		
		click(".z-menu-popup:visible .zsmenuitem-hideRow");
		Assert.assertFalse(spreadsheet.getRow(12).jq$n().isVisible());
		
		mouseDirector.openRowContextMenu(11, 13);
		click(".z-menu-popup:visible .zsmenuitem-unhideRow");
		Assert.assertTrue(spreadsheet.getRow(12).jq$n().isVisible());
	}
	
	@Test
	public void toggle_row_hide_by_drag () {
		final int ROW_6 = 5;
		spreadsheet.focus(ROW_6, 0);
		
		Header header = spreadsheet.getLeftPanel().getRowHeader(ROW_6);
		Assert.assertTrue("Header shall be visible and has width", header.isVisible() && header.getHeight() > 0);
		
		mouseDirector.dragRowToHide(ROW_6);
		Assert.assertTrue("Header shall be hidden", !header.isVisible());
		
		mouseDirector.dragRowToResize(ROW_6, 50);
		Assert.assertTrue("Header shall be unhide", header.isVisible() && header.getHeight() == 50);
		
	}
	
	@Test
	public void hide_row_by_menu_unhide_row_by_drag() {
		final int ROW_6 = 5;
		mouseDirector.openRowContextMenu(ROW_6);
		click(".z-menu-popup:visible .zsmenuitem-hideRow");
		Assert.assertFalse(spreadsheet.getRow(ROW_6).jq$n().isVisible());
		
		Header header = spreadsheet.getLeftPanel().getRowHeader(ROW_6);
		mouseDirector.dragRowToResize(ROW_6, 50);
		Assert.assertTrue("Header shall be unhide", header.isVisible() && header.getHeight() == 50);
	}
}
