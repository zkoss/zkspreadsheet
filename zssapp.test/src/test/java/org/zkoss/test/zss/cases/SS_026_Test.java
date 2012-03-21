/* SS_026_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 15, 2012 2:50:50 PM , Created by sam
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
public class SS_026_Test extends ZSSAppTest {

	@Test
	public void shift_cell_left() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration src = builder.right().build();
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-del .zstbtn-arrow");
		click(".zsmenu-deleteCell");
		click(".zsmenuitem-shiftCellLeft");
		
		verifyDelete(Delete.CELL_LEFT, src, builder, null);
	}
	
	@Test 
	public void shift_cell_up() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		CellCacheAggeration src = builder.down().build();
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-del .zstbtn-arrow");
		click(".zsmenu-deleteCell");
		click(".zsmenuitem-shiftCellUp");
		
		verifyDelete(Delete.CELL_UP, src, builder, null);	
	}
	
	@Test
	public void delete_sheet_row() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		//exam whole row is inefficient, expend 5 cells to check
		CellCacheAggeration src = builder.down().expandRight(5).build();
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-del .zstbtn-arrow");
		click(".zsmenuitem-deleteSheetRow");
		
		verifyDelete(Delete.CELL_UP, src, builder, 5);		
	}
	
	@Test
	public void delete_sheet_column() {
		int tRow = 12;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		CellCacheAggeration.Builder builder = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol);
		//exam whole row is inefficient, expend 5 cells to check
		CellCacheAggeration src = builder.right().expandDown(5).build();
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		click(".zstbtn-del .zstbtn-arrow");
		click(".zsmenuitem-deleteSheetColumn");
		
		verifyDelete(Delete.CELL_LEFT, src, builder, 5);		
	}
	

}
