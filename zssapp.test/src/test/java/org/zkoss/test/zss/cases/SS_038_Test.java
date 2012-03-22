/* SS_038_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 11:13:12 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import java.util.Iterator;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.JQuery;
import org.zkoss.test.JavascriptActions;
import org.zkoss.test.Keycode;
import org.zkoss.test.zss.Cell;
import org.zkoss.test.zss.CellCacheAggeration;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_038_Test  extends ZSSAppTest {

	@Test
	public void ctrl_v_paste_from_ctrl_x_cut() {
		int tRow = 11;
		int lCol = 8;
		int bRow = 13;
		int rCol = 9;
		keyboardDirector.ctrlCut(tRow, lCol, bRow, rCol);
		Assert.assertTrue(isVisible("div.zshighlight"));
		Assert.assertTrue(spreadsheet.isHighlight(tRow, lCol, bRow, rCol));
		
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration cutFrom = builder.build();
		spreadsheet.focus(11, 10);
		keyboardDirector.ctrlPaste(11, 10);
		
		verifyPasteAll(PasteSource.CUT, cutFrom, builder.offset(11, 10).build());
	}
	
	@Test
	public void ctrl_v_paste_from_ctrl_c_copy() {
		int tRow = 11;
		int lCol = 8;
		int bRow = 13;
		int rCol = 9;
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		Assert.assertTrue(isVisible("div.zshighlight"));
		Assert.assertTrue(spreadsheet.isHighlight(tRow, lCol, bRow, rCol));
		
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration from = builder.build();
		spreadsheet.focus(11, 10);
		keyboardDirector.ctrlPaste(11, 10);
		
		verifyPasteAll(PasteSource.COPY, from, builder.offset(11, 10).build());
	}

	@Test
	public void ctrl_d_delete_content() {
		int tRow = 11;
		int lCol = 8;
		int bRow = 13;
		int rCol = 9;
		CellCacheAggeration cache = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		keyboardDirector.ctrlD(tRow, lCol, bRow, rCol);
		
		verifyClearContent(cache);
	}
	
	@Test
	public void ctrl_delete_delete_style() {
		int tRow = 11;
		int lCol = 8;
		int bRow = 13;
		int rCol = 9;
		CellCacheAggeration cache = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		keyboardDirector.ctrlDelete(tRow, lCol, bRow, rCol);
		
		verifyClearStyle(cache);
	}

	@Test
	public void ctrl_o_openfile() {
		spreadsheet.focus(0, 0);
		
		JQuery target = spreadsheet.jq$n();
		new JavascriptActions(webDriver)
		.ctrlKeyDown(target, Keycode.O.intValue())
		.ctrlKeyUp(target, Keycode.O.intValue())
		.perform();
		
		timeBlocker.waitResponse();
		
		Assert.assertTrue(isVisible("$_openFileDialog"));
	}

	@Test
	public void ctrl_b_fontbold() {
		int tRow = 12;
		int lCol = 8;
		int bRow = 13;
		int rCol = 9;
		keyboardDirector.ctrlFontBold(tRow, lCol, bRow, rCol);
		
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			if (!c.isFontBold()) {
				System.out.println(c);
			}
			Assert.assertTrue(c.isFontBold());
		}
	}

	@Test
	public void ctrl_i_fontitalic() {
		int tRow = 12;
		int lCol = 8;
		int bRow = 13;
		int rCol = 9;
		keyboardDirector.ctrlFontItalic(tRow, lCol, bRow, rCol);
		
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertTrue(c.isFontItalic());
		}
	}
	
	@Test
	public void ctrl_i_fontunderline() {
		int tRow = 12;
		int lCol = 8;
		int bRow = 13;
		int rCol = 9;
		keyboardDirector.ctrlFontUnderline(tRow, lCol, bRow, rCol);
		
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertTrue(c.isFontUnderline());
		}
	}
	
	@Test
	public void d_clear_content() {
		int tRow = 12;
		int lCol = 8;
		int bRow = 13;
		int rCol = 9;
		CellCacheAggeration cache = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		keyboardDirector.delete(tRow, lCol, bRow, rCol);
		
		verifyClearContent(cache);
	}
	
	@Test
	public void enter_next_cell() {
		
		spreadsheet.focus(12, 10);
		JQuery K13 = getCell(12, 10).jq$n();
		
		new JavascriptActions(webDriver)
		.keyDown(K13, Keycode.ENTER.intValue())
		.keyUp(K13, Keycode.ENTER.intValue())
		.perform();
		
		Assert.assertTrue(spreadsheet.isSelection(13, 10));
	}
}
