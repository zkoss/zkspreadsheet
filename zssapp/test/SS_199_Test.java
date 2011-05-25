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
		
		//check "has header"
		click(jq("@window[title=\"Custom Sort\"] input[type=\"checkbox\"]:eq(1)"));
		waitResponse();
		
		//choose sort by "Column B"
		click(jq(" @div @combobox i.z-combobox-rounded-btn-readonly:eq(1)"));
		waitResponse();
		click(jq("@comboitem[label=\"Column B\"] td.z-comboitem-text"));
		waitResponse();
		
		click(jq("$okBtn td.z-button-cm"));
		waitResponse();
		
		//verify
		String b13value = getSpecifiedCell(1,12).text();
		String b14value = getSpecifiedCell(1,13).text();
		String b15value = getSpecifiedCell(1,14).text();
		String b16value = getSpecifiedCell(1,15).text();

		String f13value = getSpecifiedCell(5,12).text();
		String f14value = getSpecifiedCell(5,13).text();
		String f15value = getSpecifiedCell(5,14).text();
		String f16value = getSpecifiedCell(5,15).text();

		verifyEquals(b13value,"Average\u00A0total\u00A0assets");
		verifyEquals(b14value,"Current\u00A0assets");
		verifyEquals(b15value,"Fixed\u00A0assets");
		verifyEquals(b16value,"Total\u00A0assets");
		
		verifyEquals(f13value,"120,000");
		verifyEquals(f14value,"45,000");
		verifyEquals(f15value,"80,000");
		verifyEquals(f16value,"125,000");
	}
}



