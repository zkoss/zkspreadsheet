/* SS_025_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 15, 2012 12:27:18 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Test;
import org.zkoss.test.zss.CellCacheAggeration;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_025_Test extends ZSSAppTest {
	
	@Test
	public void insert_shift_cell_right() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder srcBuilder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration src = srcBuilder.build();
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-insert .zstbtn-arrow");
		click(".zsmenu-insertCell");
		click(".zsmenuitem-shiftCellRight");
		
		verifyInsert(Insert.CELL_RIGHT, src, srcBuilder);
	}
	
	@Test
	public void insert_shift_cell_down() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder srcBuilder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration src = srcBuilder.build();
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-insert .zstbtn-arrow");
		click(".zsmenu-insertCell");
		click(".zsmenuitem-shiftCellDown");
		
		verifyInsert(Insert.CELL_DOWN, src, srcBuilder);
	}
	
	@Test
	public void insert_row() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		//exam whole row will be inefficient, only expand 5 cells to check
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol)
			.expandRight(5);
		CellCacheAggeration cache = builder.build();
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-insert .zstbtn-arrow");
		click(".zsmenuitem-insertSheetRow");
		verifyInsert(Insert.CELL_DOWN, cache, builder);
	}
	
	@Test
	public void insert_column() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		//exam whole column will be inefficient, only expand 5 cells to check
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol)
			.expandDown(5);
		CellCacheAggeration cache = builder.build();
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-insert .zstbtn-arrow");
		click(".zsmenuitem-insertSheetColumn");
		verifyInsert(Insert.CELL_RIGHT, cache, builder);
	}
}
