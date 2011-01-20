/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//sort ==> Descending B13~B17
public class SS_136_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(1,12,1,16);
	
		mouseOver(jq("a.z-menu-cnt:eq(3)"));		
		waitResponse();
		click(jq("$sortDescending a.z-menu-item-cnt"));
		waitResponse();
		
		//verify
		String b13text = getSpecifiedCell(1, 12).text();
		verifyEquals(b13text, "Total assets");
		String b15text = getSpecifiedCell(1, 14).text();
		verifyEquals(b15text, "Current assets");
		String b17text = getSpecifiedCell(1, 16).text();
		verifyEquals(b17text, "Average total assets");
	}
}



