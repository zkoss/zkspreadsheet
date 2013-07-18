/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//F13:I16, custom sort, sort left to right , sort by row 13
public class SS_201_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(5,12,8,15);
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();
		
		click(jq("$sortOrientationCombo i.z-combobox-rounded-btn-readonly"));
		waitResponse();
		click(jq("@comboitem[label=\"Sort left to right\"] td.z-comboitem-text"));
		waitResponse();
		click(jq("@div @combobox i.z-combobox-rounded-btn-readonly:eq(1)"));
		waitResponse();
		click(jq("@comboitem[label=\"Row 13\"] td.z-comboitem-text"));
		waitResponse();
		click(jq("$okBtn img"));
		waitResponse();
		
		//verify
		String f13value = getSpecifiedCell(5,12).text();
		String f14value = getSpecifiedCell(5,13).text();
		String f15value = getSpecifiedCell(5,14).text();
		String f16value = getSpecifiedCell(5,15).text();

		String i13value = getSpecifiedCell(8,12).text();
		String i14value = getSpecifiedCell(8,13).text();
		String i15value = getSpecifiedCell(8,14).text();
		String i16value = getSpecifiedCell(8,15).text();

		verifyEquals(f13value,"45,000");
		verifyEquals(f14value,"80,000");
		verifyEquals(f15value,"125,000");
		verifyEquals(f16value,"122,500");

		verifyEquals(i13value,"56,000");
		verifyEquals(i14value,"80,000");
		verifyEquals(i15value,"136,000");
		verifyEquals(i16value,"128,000");
	}
}



