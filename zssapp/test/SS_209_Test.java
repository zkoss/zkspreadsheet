/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//cell K1, choose function "SUM", and input "F7" in argument 1, and then press "Tab"
public class SS_209_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {		
		rightClickCell(10,0);
		
		click(jq("$formula a.z-menu-item-cnt"));
		waitResponse();
		type(jq("$searchTextbox"), "sum");
		waitResponse();
		click(jq("$searchBtn img"));
		waitResponse();
		click(jq("@listcell[label=\"SUM\"] div.z-overflow-hidden"));
		waitResponse();
		click(jq("$okBtn img"));
		waitResponse();
		click(jq("$composeFormulaTextbox"));
		waitResponse();
		type(jq("@window[mode=\"overlapped\"][title=\"Function Arguments\"] @textbox:eq(1)"), "f7");
		waitResponse();
		keyPressNative(TAB);
		waitResponse();
		keyUpNative(TAB);
		waitResponse();
			
		//verify
		String item2 = jq(".z-listitem:eq(1)").html();
		String itemSelected = jq(".z-listitem-seld").html();		
		verifyEquals(item2, itemSelected);
	}
}



