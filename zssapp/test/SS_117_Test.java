/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/


//change row 12 height to 40
public class SS_117_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		String f13value = getSpecifiedCell(5,12).text();
		rightClickRowHeader(11);
		click(jq("$rowHeight a.z-menu-item-cnt"));
		waitResponse();
		type(jq("$headerSize"), "40");
		waitResponse();
		click(jq("$okBtn td.z-button-cm"));
		waitResponse();

		//verify, set height to 40, but expect 39 as result
		int height = getSpecifiedCell(5,11).height();
		verifyTrue(height == 39);
	}
}



