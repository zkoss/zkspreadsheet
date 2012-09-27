/* SS_028_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 15, 2012 4:35:27 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.JQuery;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_028_Test extends ZSSAppTest {

	@Test
	public void sort_ascending() {
		//B19-I22
		spreadsheet.setSelection(18, 1, 21, 8);
		
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-sortAscending");
		
		Assert.assertTrue(getCell(18, 1).getText().startsWith("A"));
		Assert.assertTrue(getCell(19, 1).getText().startsWith("C"));
		Assert.assertTrue(getCell(20, 1).getText().startsWith("O"));
		Assert.assertTrue(getCell(21, 1).getText().startsWith("T"));
	}
	
	@Test
	public void sort_descending() {
		//B19-I22
		spreadsheet.setSelection(18, 1, 21, 8);
		
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-sortDescending");
		
		Assert.assertTrue(getCell(18, 1).getText().startsWith("T"));
		Assert.assertTrue(getCell(19, 1).getText().startsWith("O"));
		Assert.assertTrue(getCell(20, 1).getText().startsWith("C"));
		Assert.assertTrue(getCell(21, 1).getText().startsWith("A"));
	}

	@Test
	public void custom_sort() {
		//B19-I22
		spreadsheet.setSelection(18, 1, 21, 8);
		
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-customSort");
		
		Assert.assertTrue("shall open customSort Dialog", isVisible("$_customSortDialog"));
		
		click(jq("$_customSortDialog $sortLevel @combobox:eq(0)').children('i.z-combobox-rounded-btn"));
		//sort by column F
		click(jq(".z-combobox-rounded-pp:visible tbody").children(":eq(5)"));
		
		click("$_customSortDialog $okBtn");
		
		Assert.assertEquals("13,750", getCell(18, 5).getText());
		Assert.assertEquals("23,000", getCell(19, 5).getText());
		Assert.assertEquals("28,000", getCell(20, 5).getText());
		Assert.assertEquals("125,000", getCell(21, 5).getText());
	}
	
	@Test
	public void toggle_apply_filter() {
		spreadsheet.focus(11, 5);//F11
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-filter");
		
		int buttonSize = jq(".zsbtn").length();
		Assert.assertEquals(6, buttonSize);
		
		//toggle off
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-filter");
		spreadsheet.focus(11, 5);
		buttonSize = jq(".zsbtn").length();
		Assert.assertEquals(0, buttonSize);
	}
	
	@Test
	public void reapply_filter() {
		//step1: input dummy data: Number, 1, 2, 3, 4, 5
		spreadsheet.focus(0, 0);
		//TODO: clear content
		//keyboardDirector.setEditText(12, 10, 17, 1, new String[]{"", "", "", "", "", "", ""});
		keyboardDirector.setEditText(12, 11, 17, 11, new String[]{"Number", "1", "2", "3", "4", "5", /*""*/});
		spreadsheet.focus(12, 11);
		
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-filter");
		
		//step2: apply autofilter with criteria, to hide "1"
		
		spreadsheet.focus(12, 11);
		JQuery autofilterBtn = jq(".zsbtn").eq(0);
		click(autofilterBtn);
		click(jq(".zsafp-item").eq(1));
		click(".zsafp-ok-btn");
		
		//step3: verify "1" is invisible
		Assert.assertFalse(spreadsheet.getRowHeader(13).jq$n().isVisible());
		
		//step4: edit "2" -> "1"
		keyboardDirector.setEditText(14, 11, "1");
		
		//step5: reapply autofilter
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-reapplyFilter");
		
		//step6: hide 2 rows
		Assert.assertFalse(spreadsheet.getRowHeader(14).jq$n().isVisible());
	}
	
	@Test
	public void clear_filter() {
		//step1: input dummy data: Number, 1, 2, 3, 4, 5
		spreadsheet.focus(0, 0);
		keyboardDirector.setEditText(12, 11, 17, 11, new String[]{"Number", "1", "2", "3", "4", "5"});
		
		//step2: apply autofilter with criteria, to hide "1"
		spreadsheet.focus(12, 11);
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-filter");
		
		//step3: apply autofilter criteria "1"
		spreadsheet.focus(12, 11);
		JQuery autofilterBtn = jq(".zsbtn").eq(0);
		click(autofilterBtn);
		click(jq(".zsafp-item").eq(1));
		click(".zsafp-ok-btn");
		
		//verify "1" is invisible
		Assert.assertFalse(spreadsheet.getRowHeader(13).jq$n().isVisible());
		
		//step4: clear autofilter
		click(".zstbtn-sortAndFilter .zstbtn-arrow");
		click(".zsmenuitem-clearFilter");
		
		//step5: verify "1" is visible
		Assert.assertTrue(spreadsheet.getRowHeader(13).jq$n().isVisible());
		
	}
}
