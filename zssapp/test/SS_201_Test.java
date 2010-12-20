/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//F13:I16, custom sort, sort left to right , sort by row 13
public class SS_201_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(5,12,8,15);
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();
		
		click(jq("$sortOrientationCombo i.z-combobox-rounded-btn-readonly"));
		waitResponse();
		click(jq("@comboitem[label=\"Sort left to right\"] td.z-comboitem-text"));
		waitResponse();
		click(jq("$sortWin @combobox i.z-combobox-rounded-btn-readonly:eq(1)"));
		waitResponse();
		click(jq("@comboitem[label=\"Row 13\"] td.z-comboitem-text"));
		waitResponse();
		click(jq("$okBtn img"));
		waitResponse();
		
		//how to verify
		sleep(5000);
	}
}



