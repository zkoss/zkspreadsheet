/* SS_041_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2012 3:05:28 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Test;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_041_Test extends ZSSAppTest {

	private void openCustomSortDialog() {
		openCustomSortDialog(12, 1, 16, 1);
	}
	
	private void openCustomSortDialog(int tRow, int lCol, int bRow, int rCol) {
		mouseDirector.openCellContextMenu(tRow, lCol, bRow, rCol);
		click(".z-menupopup:visible .zsmenu-sort");
		click(".z-menupopup:visible .zsmenuitem-customSort");
		Assert.assertTrue(isVisible("$_customSortDialog"));
	}
	
	@Test
	public void custom_sort_add_button() {
		mouseDirector.openCellContextMenu(12, 1, 16, 1);
		click(".z-menupopup:visible .zsmenu-sort");
		click(".z-menupopup:visible .zsmenuitem-customSort");
		Assert.assertTrue(isVisible("$_customSortDialog"));
		
		click("$addBtn");
		Assert.assertTrue(jq("$_customSortDialog @combobox i.z-combobox-rounded-btn-readonly:eq(3)").exists());
	}
	
	@Test
	public void custom_sort_remove_button() {
		openCustomSortDialog();
		
		click("$_customSortDialog @listcell:eq(0)");
		click("$delBtn");
		Assert.assertFalse(jq("$_customSortDialog @combobox i.z-combobox-rounded-btn-readonly:eq(1)").exists());
	}
	
	@Test
	public void custom_sort_up_arrow_button() {
		openCustomSortDialog();
		
		click("$addBtn");
		Assert.assertTrue(jq("$_customSortDialog @combobox i.z-combobox-rounded-btn-readonly:eq(3)").exists());
		
		click("$_customSortDialog @label:eq(1)");
		Assert.assertTrue(jq(".z-listitem:eq(1)").hasClass("z-listitem-seld"));
		
		click("$upBtn");
		Assert.assertTrue(jq(".z-listitem:eq(0)").hasClass("z-listitem-seld"));
	}
	
	@Test
	public void custom_sort_down_arrow_button() {
		openCustomSortDialog();
		
		click("$addBtn");
		Assert.assertTrue(jq("$_customSortDialog @combobox i.z-combobox-rounded-btn-readonly:eq(3)").exists());
		
		click("$_customSortDialog @listcell div.z-overflow-hidden:eq(0)");
		Assert.assertTrue(jq(".z-listitem:eq(0)").hasClass("z-listitem-seld"));

		click("$downBtn");
		Assert.assertTrue(jq(".z-listitem:eq(1)").hasClass("z-listitem-seld"));
	}
	
	@Test
	public void custom_sort_case_sensitive() {
		keyboardDirector.setEditText(11, 10, "A");
		keyboardDirector.setEditText(12, 10, "a");
		mouseDirector.openCellContextMenu(11, 10, 12, 10);
		click(".z-menupopup:visible .zsmenu-sort");
		click(".z-menupopup:visible .zsmenuitem-customSort");
		Assert.assertTrue(isVisible("$_customSortDialog"));
		
		jq("$_customSortDialog $caseSensitive input").getWebElement().click();
		click("$sortLevel i.z-combobox-rounded-btn:eq(0)");
		click(".z-combobox-rounded-pp:visible @comboitem[label=\"Column K\"] td.z-comboitem-text");
		click("$_customSortDialog $okBtn");
		
		Assert.assertEquals("a", getCell(11, 10).getText());
		Assert.assertEquals("A", getCell(12, 10).getText());
	}
	
	@Test
	public void custom_sort_my_data_has_headers() {
		openCustomSortDialog(11, 1, 15, 7);
		
		//check "has header"
		jq("@window[title=\"Custom Sort\"] input[type=\"checkbox\"]:eq(1)").getWebElement().click();
		
		//choose sort by "Column B"
		click(" @div @combobox i.z-combobox-rounded-btn-readonly:eq(1)");
		click("@comboitem[label=\"Column B\"] td.z-comboitem-text");
		click("$okBtn:visible");
		
		Assert.assertEquals("Average total assets", getCell(12, 1).getEdit());
		Assert.assertEquals("Current assets", getCell(13, 1).getEdit());
		Assert.assertEquals("Fixed assets", getCell(14, 1).getEdit());
		Assert.assertEquals("Total assets", getCell(15, 1).getEdit());
		
		Assert.assertEquals("120,000", getCell(12, 5).getText());
		Assert.assertEquals("45,000", getCell(13, 5).getText());
		Assert.assertEquals("80,000", getCell(14, 5).getText());
		Assert.assertEquals("125,000", getCell(15, 5).getText());
	}
	
	@Test
	public void custom_sort_sort_top_to_bottom() {
		openCustomSortDialog(12, 1, 15, 7);
		
		//choose sort by "Column B"
		click(" @div @combobox i.z-combobox-rounded-btn-readonly:eq(1)");
		click("@comboitem[label=\"Column B\"] td.z-comboitem-text");
		click("$okBtn:visible");
		
		Assert.assertEquals("Average total assets", getCell(12, 1).getEdit());
		Assert.assertEquals("Current assets", getCell(13, 1).getEdit());
		Assert.assertEquals("Fixed assets", getCell(14, 1).getEdit());
		Assert.assertEquals("Total assets", getCell(15, 1).getEdit());
		
		Assert.assertEquals("120,000", getCell(12, 5).getText());
		Assert.assertEquals("45,000", getCell(13, 5).getText());
		Assert.assertEquals("80,000", getCell(14, 5).getText());
		Assert.assertEquals("125,000", getCell(15, 5).getText());
	}
	
	@Test
	public void custom_sort_sort_left_to_right() {
		openCustomSortDialog(12, 5, 15, 8);
		
		//choose sort by "Column B"
		click("$sortOrientationCombo i.z-combobox-rounded-btn-readonly");
		click("@comboitem[label=\"Sort left to right\"] td.z-comboitem-text");
		click("@div @combobox i.z-combobox-rounded-btn-readonly:eq(1)");
		click("@comboitem[label=\"Row 13\"] td.z-comboitem-text");
		click("$okBtn");
		
		Assert.assertEquals("45,000", getCell(12, 5).getText());
		Assert.assertEquals("80,000", getCell(13, 5).getText());
		Assert.assertEquals("125,000", getCell(14, 5).getText());
		Assert.assertEquals("122,500", getCell(15, 5).getText());

		Assert.assertEquals("56,000", getCell(12, 8).getText());
		Assert.assertEquals("80,000", getCell(13, 8).getText());
		Assert.assertEquals("136,000", getCell(14, 8).getText());
		Assert.assertEquals("128,000", getCell(15, 8).getText());
	}
	
	@Test
	public void custom_sort_sort_by() {
		openCustomSortDialog(12, 1, 15, 7);
		
		click(" @div @combobox i.z-combobox-rounded-btn-readonly:eq(2)");
		click("@comboitem[label=\"From Z to A\"] td.z-comboitem-text");
		click("@div @combobox i.z-combobox-rounded-btn-readonly:eq(1)");
		click("@comboitem[label=\"Column B\"] td.z-comboitem-text");
		click("$okBtn");
		
		Assert.assertEquals("#VALUE!", getCell(12, 6).getText());
		Assert.assertEquals("80,000", getCell(13, 6).getText());
		Assert.assertEquals("46,000", getCell(14, 6).getText());
		Assert.assertEquals("83,000", getCell(15, 6).getText());
				
		Assert.assertEquals("#VALUE!", getCell(12, 5).getText());
		Assert.assertEquals("80,000", getCell(13, 5).getText());
		Assert.assertEquals("45,000", getCell(14, 5).getText());
		Assert.assertEquals("82,500", getCell(15, 5).getText());
	}
}
