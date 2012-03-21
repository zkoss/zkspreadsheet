/* SS_027_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 15, 2012 3:45:01 PM , Created by sam
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
public class SS_027_Test extends ZSSAppTest {

	@Test
	public void clear_content() {
		int tRow = 11;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		CellCacheAggeration cache = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		
		click(".zstbtn-clear .zstbtn-arrow");
		click(".zsmenuitem-clearContent");
		
		verifyClearContent(cache);
	}
	
	@Test
	public void clear_style() {
		int tRow = 11;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		CellCacheAggeration cache = getCellCacheAggerationBuilder(tRow, lCol, bRow, rCol).build();
		
		click(".zstbtn-clear .zstbtn-arrow");
		click(".zsmenuitem-clearStyle");
		
		verifyClearStyle(cache);
	}
	
	@Test
	public void clear_all() {
		int tRow = 11;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		
		click(".zstbtn-clear .zstbtn-arrow");
		click(".zsmenuitem-clearAll");
		
		verifyClearAll(tRow, lCol, bRow, rCol);
	}
	

}
