/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//B13:G16, custom sort, sort by Z to A , sort by column B
public class SS_203_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(1,12,7,15);
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();
		
		click(jq("$sortWin @combobox i.z-combobox-rounded-btn-readonly:eq(2)"));
		waitResponse();
		click(jq("@comboitem[label=\"From Z to A\"] td.z-comboitem-text"));
		waitResponse();
		
		click(jq("$sortWin @combobox i.z-combobox-rounded-btn-readonly:eq(1)"));
		waitResponse();
		click(jq("@comboitem[label=\"Column B\"] td.z-comboitem-text"));
		waitResponse();
		click(jq("$okBtn td.z-button-cm"));
		waitResponse();
		
		//verify
		String g13value = getSpecifiedCell(6,12).text();
		String g14value = getSpecifiedCell(6,13).text();
		String g15value = getSpecifiedCell(6,14).text();
		String g16value = getSpecifiedCell(6,15).text();

		String f13value = getSpecifiedCell(5,12).text();
		String f14value = getSpecifiedCell(5,13).text();
		String f15value = getSpecifiedCell(5,14).text();
		String f16value = getSpecifiedCell(5,15).text();

		verifyEquals(g13value,"#VALUE!");
		verifyEquals(g14value,"80,000");
		verifyEquals(g15value,"46,000");
		verifyEquals(g16value,"83,000");
				
		verifyEquals(f13value,"#VALUE!");
		verifyEquals(f14value,"80,000");
		verifyEquals(f15value,"45,000");
		verifyEquals(f16value,"82,500");
	}
}



