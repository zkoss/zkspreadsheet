/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//B13:G16, custom sort, sort top to bottom , sort by column B
public class SS_200_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(1,12,7,15);
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$customSort a.z-menu-item-cnt"));
		waitResponse();
		
		click(jq(" @div @combobox i.z-combobox-rounded-btn-readonly:eq(1)"));
		waitResponse();
		click(jq("@comboitem[label=\"Column B\"] td.z-comboitem-text"));
		waitResponse();
		click("$okBtn:visible td.z-button-cm");
		waitResponse();
		
		int nbsp = 160;
		String aChar = new Character((char)nbsp).toString();

		//verify
		String b13value = getCellText(1,12);
		String b14value = getCellText(1,13);
		String b15value = getCellText(1,14);
		String b16value = getCellText(1,15);

		String f13value = getCellText(5,12);
		String f14value = getCellText(5,13);
		String f15value = getCellText(5,14);
		String f16value = getCellText(5,15);

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



