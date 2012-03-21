/* SS_024_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 15, 2012 11:42:37 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import java.util.Iterator;

import org.junit.Assert;
import org.junit.Test;
import org.zkoss.test.zss.Cell;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_024_Test extends ZSSAppTest {

	@Test
	public void merge_and_center_by_toolbarbutton() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-mergeAndCenter");
		
		Cell cell = getCell(tRow, lCol);
		verifyMerge(true, tRow, lCol, bRow, rCol);
		Assert.assertEquals("center", cell.getHorizontalAlign());
	}
	
	@Test
	public void merge_and_center_by_menuitem() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-mergeAndCenter .zstbtn-arrow");
		click(".zsmenuitem-mergeAndCenter");
		
		Cell cell = getCell(tRow, lCol);
		verifyMerge(true, tRow, lCol, bRow, rCol);
		Assert.assertEquals("center", cell.getHorizontalAlign());
	}
	
	@Test
	public void merge_across() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-mergeAndCenter .zstbtn-arrow");
		click(".zsmenuitem-mergeAcross");
		
		Iterator<Cell> i = iterator(tRow, lCol, bRow, lCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertTrue(c.isMerged());
		}
	}
	
	@Test
	public void merge_cell() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-mergeAndCenter .zstbtn-arrow");
		click(".zsmenuitem-mergeCell");
		
		verifyMerge(true, tRow, lCol, bRow, rCol);
	}
	
	@Test
	public void unmerge_cell() {
		int tRow = 11;
		int lCol = 1;
		int bRow = 11;
		int rCol = 4;
		spreadsheet.focus(tRow, lCol);
		click(".zstbtn-mergeAndCenter .zstbtn-arrow");
		click(".zsmenuitem-unmergeCell");
		
		verifyMerge(false, tRow, lCol, bRow, rCol);
	}
	
	void verifyMerge(boolean shallMerge, int tRow, int lCol, int bRow, int rCol) {
		Iterator<Cell> i = iterator(tRow, lCol, bRow, rCol);
		while (i.hasNext()) {
			Cell c = i.next();
			Assert.assertEquals(shallMerge, c.isMerged());
		}
	}
}
