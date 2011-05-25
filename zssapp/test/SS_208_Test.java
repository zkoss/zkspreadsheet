/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//cell K1, choose function "SUM", and input "F7,F8,F9" as formula textbox, click "OK"
public class SS_208_Test extends SSAbstractTestCase {
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
		type(jq("$composeFormulaTextbox"), "f7,f8");
		waitResponse();
		click(jq("@window[title=\"Function Arguments\"] $okBtn img"));
		waitResponse();
		
		//verify
		String k1value = getSpecifiedCell(10,0).text();
		verifyEquals(k1value, "161500");
	}
}



