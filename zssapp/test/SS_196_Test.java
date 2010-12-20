/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//sort ==> custom sort B13~B17, add second condition and move it upward to be first condition
public class SS_196_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(1,12,1,16);
	
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();

		//verify
		String titleOfPopup =  jq(".z-window-highlighted.z-window-highlighted-shadow .z-window-highlighted-header").attr("textContent");
		verifyEquals(titleOfPopup,"Custom Sort");		
		verifyFalse(jq("$sortWin @combobox i.z-combobox-rounded-btn-readonly:eq(3)").exists());
		
		click(jq("$addBtn img"));
		waitResponse();

		//verify
		verifyTrue(jq("$sortWin @combobox i.z-combobox-rounded-btn-readonly:eq(3)").exists());

		click(jq("$sortWin @label:eq(1)"));
		
		//verify
		verifyTrue(jq(".z-listitem:eq(1)").hasClass("z-listitem-seld"));
		
		waitResponse();
		click(jq("$upBtn img"));
		waitResponse();
		
		//verify
		verifyTrue(jq(".z-listitem:eq(0)").hasClass("z-listitem-seld"));
	}
}



