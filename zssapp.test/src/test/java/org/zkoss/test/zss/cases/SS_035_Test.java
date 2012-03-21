/* SS_035_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 16, 2012 1:57:53 PM , Created by sam
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
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase 
public class SS_035_Test extends ZSSAppTest {
	
	@Test
	public void column_font_family() {
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zsfontfamily .z-combobox-btn");
		click(".zsfontfamily-arial");
		
		verifyFontFamily("arial", 5, 5, 20, 5);
	}
	
	@Test
	public void column_font_size() {
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zsfontsize .z-combobox-btn");
		click(".zsfontsize-20");
		
		verifyFontSize(20, 5, 5, 20, 5);
	}
	
	@Test
	public void column_font_bold() {
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zstbtn-fontBold");
		
		verifyFontBold(true, 5, 5, 20, 5);
	}
	
	@Test
	public void column_font_italic() {
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zstbtn-fontItalic");
		
		verifyFontItalic(true, 5, 5, 20, 5);
	}
	
	@Test
	public void column_font_color() {
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zstbtn-fontColor .zstbtn-cave");
		JQuery target = jqFactory.create("'.z-colorpalette-colorbox'").eq(68);
		String color = target.css("background-color");
		click(target);
		
		verifyFontColor(new Color(color), 5, 5, 20, 5);
	}
	
	@Test
	public void column_fill_color() {
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zstbtn-fillColor .zstbtn-cave");
		JQuery target = jqFactory.create("'.z-colorpalette-colorbox'").eq(68);
		String color = target.css("background-color");
		click(target);
		
		verifyFillColor(new Color(color), 5, 5, 20, 5);
	}
	
	@Test
	public void column_border() {
		int tRow = 5;
		int lCol = 5;
		int bRow = 20;
		int rCol = 5;
		CellCacheAggeration cache = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zstbtn-border .zstbtn-arrow");
		click(".zsmenuitem-" + BorderType.LEFT.toString());
		
		verifyBorder(BorderType.LEFT, "#000000", tRow, lCol, bRow, rCol, cache);
	}
	
	@Test
	public void column_vertical_align() {
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zstbtn-verticalAlign .zstbtn-arrow");
		click(".zsmenuitem-" + AlignType.VERTICAL_ALIGN_TOP.toString());
		
		verifyAlign(AlignType.VERTICAL_ALIGN_TOP, 5, 5, 20, 5);
	}
	
	@Test
	public void column_horizontal_align() {
		mouseDirector.openColumnContextMenu(5);
		
		click(".zsstylepanel .zstbtn-horizontalAlign .zstbtn-arrow");
		click(".zsmenuitem-" + AlignType.HORIZONTAL_ALIGN_LEFT.toString());
		
		verifyAlign(AlignType.HORIZONTAL_ALIGN_LEFT, 5, 5, 20, 5);
	}
	
	@Test
	public void column_cut() {
		CellCacheAggeration from = getCellCacheAggerationBuilder(5, 5, 20, 5).build();
		
		mouseDirector.openColumnContextMenu(5);//column F
		click(".z-menu-popup:visible .zsmenuitem-cut");
		
		mouseDirector.openColumnContextMenu(10);//column K
		click(".z-menu-popup:visible .zsmenuitem-paste");
		
		verifyPasteAll(PasteSource.CUT, from, 5, 10, 20, 10);
	}
	
	@Test
	public void column_copy() {
		CellCacheAggeration from = getCellCacheAggerationBuilder(5, 5, 20, 5).build();
		
		mouseDirector.openColumnContextMenu(5);//column F
		click(".z-menu-popup:visible .zsmenuitem-copy");
		
		mouseDirector.openColumnContextMenu(10);//column K
		click(".z-menu-popup:visible .zsmenuitem-paste");
		
		verifyPasteAll(PasteSource.COPY, from, 5, 10, 20, 10);
	}
	
	
	@Test
	public void column_paste_special() {
		CellCacheAggeration from = getCellCacheAggerationBuilder(5, 5, 20, 5).build();
		mouseDirector.openColumnContextMenu(5);//column F
		click(".z-menu-popup:visible .zsmenuitem-copy");
		
		mouseDirector.openColumnContextMenu(10);//column K
		click(".z-menu-popup:visible .zsmenuitem-pasteSpecial");
		Assert.assertTrue(isVisible("$_pasteSpecialDialog"));
		click("$_pasteSpecialDialog $okBtn");
		
		verifyPasteAll(PasteSource.COPY, from, 5, 10, 20, 10);
	}
	
	@Test
	public void column_insert() {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(5, 5, 20, 5);
		CellCacheAggeration from = builder.build();
		mouseDirector.openColumnContextMenu(5);//column F
		
		click(".z-menu-popup:visible .zsmenuitem-insertSheetColumn");
		
		verifyInsert(Insert.CELL_RIGHT, from, builder);
	}
	
	@Test
	public void column_delete() {
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(5, 5, 20, 5);
		CellCacheAggeration from = builder.right().build();
		mouseDirector.openColumnContextMenu(5);//column F
		
		click(".z-menu-popup:visible .zsmenuitem-deleteSheetColumn");
		
		verifyDelete(Delete.CELL_LEFT, from, builder, null);
	}
	
	@Test
	public void column_clear_content() {
		CellCacheAggeration cache = getCellCacheAggerationBuilder(5, 5, 20, 5).build();
		mouseDirector.openColumnContextMenu(5);//column F
		click(".z-menu-popup:visible .zsmenuitem-clearContent");
		verifyClearContent(cache);
	}
	
	@Test
	public void column_format_number() {
		
		mouseDirector.openColumnContextMenu(5);//column F
		click(".z-menu-popup:visible .zsmenuitem-formatCell");
		Assert.assertTrue(isVisible("$_formatNumberDialog"));
		
		click("$format a.z-menu-item-cnt");
		click("@listcell[label=\"Accounting\"] div.z-overflow-hidden");
		click("@listcell[label=\"$1,234\"] div.z-overflow-hidden");
		click("$_formatNumberDialog $okBtn");
		
		verifyFormatNumber(FormatNumber.ACCOUNTING, 5, 5, 20, 5);
	}
	
	@Test
	public void column_width() {
		mouseDirector.openColumnContextMenu(5);//column F
		
		click(".z-menu-popup:visible .zsmenuitem-columnWidth");
		WebElement input =  jq("$headerSize").getWebElement();
		input.clear();
		input.sendKeys("120");
		input.sendKeys(Keys.ENTER);
		timeBlocker.waitUntil(1);
		timeBlocker.waitResponse();
		
		Assert.assertEquals(115, getCell(11, 5).jq$n().width());
	}
	
	@Test
	public void toggle_column_hide() {
		mouseDirector.openColumnContextMenu(5);//column F
		
		click(".z-menu-popup:visible .zsmenuitem-hideColumn");
		Assert.assertFalse(spreadsheet.getColumn(5).jq$n().isVisible());
		
		mouseDirector.openColumnContextMenu(4, 6);
		click(".z-menu-popup:visible .zsmenuitem-unhideColumn");
		Assert.assertTrue(spreadsheet.getColumn(5).jq$n().isVisible());
	}
}
