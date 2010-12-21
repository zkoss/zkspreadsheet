/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//cell K1, find function in category "Text"
public class SS_205_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {		
		rightClickCell(10,0);
		
		click(jq("$formula a.z-menu-item-cnt"));
		waitResponse();
		click(jq("$categoryCombobox i.z-combobox-btn"));
		waitResponse();
		click(jq("@comboitem[label=\"Text\"] td.z-comboitem-text"));
		waitResponse();
		
		//verify
		String func1 = jq(".z-listcell-cnt:eq(0)").text();
		verifyEquals(func1, "ASC");
		String func2 = jq(".z-listcell-cnt:eq(1)").text();
		verifyEquals(func2, "BAHTTEXT");
		String func6 = jq(".z-listcell-cnt:eq(5)").text();
		verifyEquals(func6, "CONCATENATE");
	}
}



