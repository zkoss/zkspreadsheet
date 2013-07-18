/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//sort ==> custom sort I1,I2, case sensitive
public class SS_198_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(8, 0, 8, 0);
		type(jq("$formulaEditor"), "AAA");
		waitResponse();
		selectCells(8, 1, 8, 1);
		type(jq("$formulaEditor"), "aaa");
		waitResponse();
	
		rightClickCells(8,0,8,1);
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();
		
		//check "Case sensitive"
		click(jq("$_customSortDialog $caseSensitive"));
		waitResponse();
		
		//choose sort by "Column I"
		click(jq("$sortLevel i.z-combobox-rounded-btn:eq(0)"));
		waitResponse();
		click(jq(".z-combobox-rounded-pp:visible @comboitem[label=\"Column I\"] td.z-comboitem-text"));
		waitResponse();
		
		//click OK button
		click(jq("$okBtn td.z-button-cm"));
		waitResponse();
		
		//how to verify
		String i1value = getSpecifiedCell(8,0).text();
		String i2value = getSpecifiedCell(8,1).text();
		verifyEquals(i1value,"aaa");
		verifyEquals(i2value,"AAA");	
	}
}



