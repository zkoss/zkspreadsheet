/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//B12:G16, custom sort, has headers, sort by column B
public class SS_199_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(1,11,7,15);
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();
		
		click(jq("input[type=\"checkbox\"]:eq(2)"));
		waitResponse();
		click(jq("$sortWin @combobox i.z-combobox-rounded-btn-readonly:eq(1)"));
		waitResponse();
		click(jq("@comboitem[label=\"Column B\"] td.z-comboitem-text"));
		waitResponse();
		click(jq("$okBtn td.z-button-cm"));
		waitResponse();
		
		//how to verify
		sleep(5000);
	}
}



