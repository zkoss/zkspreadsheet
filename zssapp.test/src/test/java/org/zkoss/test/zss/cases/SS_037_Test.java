/* SS_037_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 16, 2012 5:39:31 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
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
public class SS_037_Test extends ZSSAppTest {
	
	@Test
	public void cell_font_family() {
		mouseDirector.openCellContextMenu(12, 5, 15, 8);
		
		click(".zsstylepanel .zsfontfamily .z-combobox-btn");
		click(".zsfontfamily-arial");
		
		verifyFontFamily("arial", 12, 5, 15, 8);
	}
	
	@Test
	public void cell_font_size() {
		mouseDirector.openCellContextMenu(12, 5, 15, 8);
		
		click(".zsstylepanel .zsfontsize .z-combobox-btn");
		click(".zsfontsize-20");
		
		verifyFontSize(20, 12, 5, 15, 8);
	}
	
	@Test
	public void cell_font_bold() {
		mouseDirector.openCellContextMenu(12, 5, 15, 8);
		
		click(".zsstylepanel .zstbtn-fontBold");
		
		verifyFontBold(true, 12, 5, 15, 8);
	}
	@Test
	public void cell_font_italic() {
		mouseDirector.openCellContextMenu(12, 5, 15, 8);
		
		click(".zsstylepanel .zstbtn-fontItalic");
		
		verifyFontItalic(true, 12, 5, 15, 8);
	}
	
	@Test
	public void cell_font_color() {
		mouseDirector.openCellContextMenu(12, 5, 15, 8);
		
		click(".zsstylepanel .zstbtn-fontColor .zstbtn-cave");
		JQuery target = jqFactory.create("'.z-colorpalette-colorbox'").eq(68);
		String color = target.css("background-color");
		click(target);
		
		verifyFontColor(new Color(color), 12, 5, 15, 8);
	}
	
	@Test
	public void cell_fill_color() {
		mouseDirector.openCellContextMenu(12, 5, 15, 8);
		
		click(".zsstylepanel .zstbtn-fillColor .zstbtn-cave");
		JQuery target = jqFactory.create("'.z-colorpalette-colorbox'").eq(68);
		String color = target.css("background-color");
		click(target);
		
		verifyFillColor(new Color(color), 12, 5, 15, 8);
	}
	
	@Test
	public void cell_border() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 15;
		int rCol = 8;
		CellCacheAggeration cache = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		
		click(".zsstylepanel .zstbtn-border .zstbtn-arrow");
		click(".zsmenuitem-" + BorderType.ALL.toString());
		
		verifyBorder(BorderType.ALL, "#000000", tRow, lCol, bRow, rCol, cache);
	}
	
	@Test
	public void cell_vertical_align() {
		mouseDirector.openCellContextMenu(12, 5, 15, 8);
		
		click(".zsstylepanel .zstbtn-verticalAlign .zstbtn-arrow");
		click(".zsmenuitem-" + AlignType.VERTICAL_ALIGN_TOP.toString());
		
		verifyAlign(AlignType.VERTICAL_ALIGN_TOP, 12, 5, 15, 8);
	}
	
	@Test
	public void cell_horizontal_align() {
		mouseDirector.openCellContextMenu(12, 5, 15, 8);
		
		click(".zsstylepanel .zstbtn-horizontalAlign .zstbtn-arrow");
		click(".zsmenuitem-" + AlignType.HORIZONTAL_ALIGN_LEFT.toString());
		
		verifyAlign(AlignType.HORIZONTAL_ALIGN_LEFT, 12, 5, 15, 8);
	}
	
	@Test
	public void cell_paste_from_cut() {
		int tRow = 11;
		int lCol = 5;
		int bRow = 16;
		int rCol = 5;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenuitem-cut");
		
		spreadsheet.focus(11, 10);
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-paste");
		
		verifyPasteAll(PasteSource.CUT, copyFrom, builder.offset(11, 10).build());
	}
	
	@Test
	public void cell_paste_from_copy() {
		int tRow = 11;
		int lCol = 5;
		int bRow = 16;
		int rCol = 5;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenuitem-copy");
		
		spreadsheet.focus(11, 10);
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-paste");
		
		verifyPasteAll(PasteSource.COPY, copyFrom, builder.offset(11, 10).build());
	}
	
	@Test
	public void cell_paste_special() {
		int tRow = 11;
		int lCol = 5;
		int bRow = 16;
		int rCol = 5;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration copyFrom = builder.build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenuitem-copy");
		
		spreadsheet.focus(11, 10);
		mouseDirector.openCellContextMenu(11, 10);
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		
		Assert.assertTrue(isVisible("$_pasteSpecialDialog"));
		click("$_pasteSpecialDialog $okBtn");
		Assert.assertEquals(copyFrom, builder.offset(11, 10).build());
	}
	
	@Test
	public void cell_insert_right() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration src = builder.build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenu-insert");
		click(".z-menu-popup:visible .zsmenuitem-shiftCellRight");
		
		verifyInsert(Insert.CELL_RIGHT, src, builder);
	}
	
	@Test
	public void cell_insert_down() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration src = builder.build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenu-insert");
		click(".z-menu-popup:visible .zsmenuitem-shiftCellDown");
		
		verifyInsert(Insert.CELL_DOWN, src, builder);
	}
	
	@Test
	public void cell_insert_row() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		//exam whole row will be inefficient, only expand 5 cells to check
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol)
			.expandRight(5);
		CellCacheAggeration cache = builder.build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenu-insert");
		click(".z-menu-popup:visible .zsmenuitem-insertSheetRow");
		
		verifyInsert(Insert.CELL_DOWN, cache, builder);
	}
	
	@Test
	public void cell_insert_column() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		//exam whole column will be inefficient, only expand 5 cells to check
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol)
			.expandDown(5);
		CellCacheAggeration cache = builder.build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenu-insert");
		click(".z-menu-popup:visible .zsmenuitem-insertSheetColumn");
		verifyInsert(Insert.CELL_RIGHT, cache, builder);
	}
	
	@Test
	public void cell_delete_left() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration src = builder.right().build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenu-del");
		click(".z-menu-popup:visible .zsmenuitem-shiftCellLeft");
		
		verifyDelete(Delete.CELL_LEFT, src, builder, null);
	}
	
	@Test
	public void cell_delete_up() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration src = builder.down().build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenu-del");
		click(".z-menu-popup:visible .zsmenuitem-shiftCellUp");
		
		verifyDelete(Delete.CELL_UP, src, builder, null);	
	}
	
	@Test
	public void cell_delete_row() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		//exam whole row is inefficient, expend 5 cells to check
		CellCacheAggeration src = builder.down().expandRight(5).build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenu-del");
		click(".z-menu-popup:visible .zsmenuitem-deleteSheetRow");
		
		verifyDelete(Delete.CELL_UP, src, builder, 5);	
	}
	
	@Test
	public void cell_delete_column() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		//exam whole row is inefficient, expend 5 cells to check
		CellCacheAggeration src = builder.right().expandDown(5).build();
		
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenu-del");
		click(".z-menu-popup:visible .zsmenuitem-deleteSheetColumn");
		
		verifyDelete(Delete.CELL_LEFT, src, builder, 5);	
	}
	
	@Test
	public void cell_clear_content() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration src = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menu-popup:visible .zsmenuitem-clearContent");
		
		verifyClearContent(src);
	}
	
	@Test
	public void cell_filter_filter_reapply() {
		int tRow = 12;
		int lCol = 11;
		int bRow = 17;
		int rCol = 11;
		keyboardDirector.setEditText(tRow, lCol, bRow, rCol, new String[]{"Number", "1", "2", "3", "4", "5", /*""*/});
		mouseDirector.openCellContextMenu(tRow, lCol);
		
		//step1: apply filter
		click(".z-menu-popup:visible .zsmenu-filter");
		click(".z-menu-popup:visible .zsmenuitem-filter");
		
		//step2: apply autofilter with criteria, to hide "1"
		spreadsheet.focus(tRow, lCol);
		JQuery autofilterBtn = jq(".zsbtn").eq(0);
		click(autofilterBtn);
		click(jq(".zsafp-item").eq(1));
		click(".zsafp-ok-btn");
		
		//step3: verify "1" is invisible
		Assert.assertFalse(spreadsheet.getRowHeader(13).jq$n().isVisible());
		
		//step4: edit "2" -> "1"
		keyboardDirector.setEditText(14, 11, "1");
		
		//step5: reapply autofilter
		mouseDirector.openCellContextMenu(tRow, lCol);
		click(".z-menu-popup:visible .zsmenu-filter");
		click(".z-menu-popup:visible .zsmenuitem-reapplyFilter");
		
		//step6: hide 2 rows
		Assert.assertFalse(spreadsheet.getRowHeader(14).jq$n().isVisible());
	}
	
	@Test
	public void cell_sort_ascending() {
		//B19-I22
		mouseDirector.openCellContextMenu(18, 1, 21, 8);
		
		click(".z-menu-popup:visible .zsmenu-sort");
		click(".z-menu-popup:visible .zsmenuitem-sortAscending");
		
		Assert.assertTrue(getCell(18, 1).getText().startsWith("A"));
		Assert.assertTrue(getCell(19, 1).getText().startsWith("C"));
		Assert.assertTrue(getCell(20, 1).getText().startsWith("O"));
		Assert.assertTrue(getCell(21, 1).getText().startsWith("T"));
	}
	
	@Test
	public void cell_sort_descending() {
		mouseDirector.openCellContextMenu(18, 1, 21, 8);
		
		click(".z-menu-popup:visible .zsmenu-sort");
		click(".z-menu-popup:visible .zsmenuitem-sortDescending");
		
		Assert.assertTrue(getCell(18, 1).getText().startsWith("T"));
		Assert.assertTrue(getCell(19, 1).getText().startsWith("O"));
		Assert.assertTrue(getCell(20, 1).getText().startsWith("C"));
		Assert.assertTrue(getCell(21, 1).getText().startsWith("A"));
	}
	
	@Test
	public void cell_custom_sort() {
		mouseDirector.openCellContextMenu(18, 1, 21, 8);
		
		click(".z-menu-popup:visible .zsmenu-sort");
		click(".z-menu-popup:visible .zsmenuitem-customSort");
		
		Assert.assertTrue("shall open customSort Dialog", isVisible("$_customSortDialog"));
		
		jq("$_customSortDialog $sortLevel @combobox:eq(0)').children('i.z-combobox-rounded-btn").getWebElement().click();
		//sort by column F
		jq(".z-combobox-rounded-pp:visible tbody").children().eq(5).getWebElement().click();
		
		click("$_customSortDialog $okBtn");
		
		Assert.assertEquals("13,750", getCell(18, 5).getText());
		Assert.assertEquals("23,000", getCell(19, 5).getText());
		Assert.assertEquals("28,000", getCell(20, 5).getText());
		Assert.assertEquals("125,000", getCell(21, 5).getText());	
	}
	
	@Test
	public void cell_hyperlink() {
		mouseDirector.openCellContextMenu(11, 10);
		
		click(".z-menu-popup:visible .zsmenuitem-hyperlink");
		Assert.assertTrue(isVisible("$_insertHyperlinkDialog"));
		
        WebElement inp = jq("$addrCombobox input.z-combobox-inp").getWebElement();
        inp.sendKeys("http://ja.wikipedia.org/wiki");
        inp.sendKeys(Keys.TAB);
        timeBlocker.waitUntil(2);
        click("$_insertHyperlinkDialog $okBtn");
        timeBlocker.waitResponse();
        
        JQuery link = getCell(11, 10).jq$n("real").children().first();
        
        Assert.assertEquals("http://ja.wikipedia.org/wiki", link.attr("href"));
	}
}
