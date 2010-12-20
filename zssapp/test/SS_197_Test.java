/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//sort ==> custom sort B13~B17, add second condition and downward first condition to the bottom 
public class SS_197_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(1,12,1,16);
	
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();
		
		click(jq("$addBtn img"));
		waitResponse();
		click(jq("$sortWin @listcell div.z-overflow-hidden:eq(0)"));
		waitResponse();
		click(jq("$downBtn img"));
		waitResponse();
		
		//how to verify
		sleep(5000);
	}
}



