/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//sort ==> custom sort B13~B17, add second condition
public class SS_194_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(1,12,1,16);
	
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();
		
		//verify
		verifyTrue(jq("$_customSortDialog").isVisible());		

		//verify
		verifyFalse(jq("$_customSortDialog @combobox i.z-combobox-rounded-btn-readonly:eq(3)").exists());
		
		click(jq("$addBtn td.z-button-bm"));
		waitResponse();
		
		//verify
		verifyTrue(jq("$_customSortDialog @combobox i.z-combobox-rounded-btn-readonly:eq(3)").exists());
		
	}
}



