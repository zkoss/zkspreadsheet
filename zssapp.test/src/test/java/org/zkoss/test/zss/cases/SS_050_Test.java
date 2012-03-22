/* SS_240_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 8, 2012 9:21:17 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.JQuery;
import org.zkoss.test.zss.Rect;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * Select sheet test
 * 
 * @author sam
 *
 */
@ZSSTestCase
public class SS_050_Test extends ZSSAppTest {

	@Test
	public void select_sheet_focus() {
		//steps
		//1. focus [5, 10] at first sheet
		//2. switch back to first sheet, focus shall remain [5, 10]
		
		spreadsheet.focus(5, 10);
		
		JQuery secondSheet = jq(".zssheettab").eq(1);
		click(secondSheet);
		
		JQuery firstSheet = jq(".zssheettab").first();
		click(firstSheet);
		Assert.assertTrue(spreadsheet.isSelection(5, 10));
	}
		
	@Test
	public void select_sheet_visible_range() {
		//steps
		//1. scroll down
		//2. select second sheet
		//3. switch back to first sheet, shall restore previous top/left
		mouseDirector.pageDown(0, 0);
		Rect expected = spreadsheet.getVisibleRange();
		Assert.assertTrue(expected.getTop() > 0);
		
		JQuery secondSheet = jq(".zssheettab").eq(1);
		click(secondSheet);
		
		JQuery firstSheet = jq(".zssheettab").first();
		click(firstSheet);
		
		Assert.assertEquals(expected, spreadsheet.getVisibleRange());
	}
	
	
	@Test
	public void select_sheet_highlight() {
		int tRow = 11;
		int lCol = 5;
		int bRow = 18;
		int rCol = 8;
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		keyboardDirector.ctrlCopy(tRow, lCol, bRow, rCol);
		Assert.assertTrue(isVisible(".zshighlight"));
		
		JQuery secondSheet = jq(".zssheettab").eq(1);
		click(secondSheet);
		Assert.assertFalse(isVisible(".zshighlight"));
		
		JQuery firstSheet = jq(".zssheettab").first();
		click(firstSheet);
		Assert.assertTrue(isVisible(".zshighlight"));
		Assert.assertTrue(spreadsheet.isHighlight(tRow, lCol, bRow, rCol));
	}
	
	@Test
	public void select_sheet_autofilter() {
		
		Assert.assertEquals(0, jq(".zsdropdown").length());
		
		spreadsheet.focus(12, 5);
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-filter");
		Assert.assertTrue(isVisible(".zsdropdown"));
		Assert.assertTrue(jq(".zsdropdown").length() > 0);
		
		JQuery secondSheet = jq(".zssheettab").eq(1);
		click(secondSheet);
		Assert.assertEquals(0, jq(".zsdropdown").length());
		Assert.assertFalse(isVisible(".zsdropdown"));
		
		JQuery firstSheet = jq(".zssheettab").first();
		click(firstSheet);
		Assert.assertTrue(isVisible(".zsdropdown"));
		Assert.assertTrue(jq(".zsdropdown").length() > 0);
		
		//switch back again, shall not have autofilter
		click(secondSheet);
		Assert.assertEquals(0, jq(".zsdropdown").length());
		Assert.assertFalse(isVisible(".zsdropdown"));
	}
	
//	//TODO: when zssapp support add data validation
//	@Test
//	public void select_sheet_data_validation() {
//		
//	}
}
