/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//change column F's width to 120
public class SS_108_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickColumnHeader(5);		
		click(jq("$columnWidth a.z-menu-item-cnt"));
		waitResponse();
		type(jq("$headerSize"), "120");
		waitResponse();
		click(jq("$okBtn td.z-button-cm"));
		waitResponse();

		//verify, set width to 120, but expect 115 as result?
		int width = getSpecifiedCell(5,11).width();
		verifyTrue(width == 115);
	}
}