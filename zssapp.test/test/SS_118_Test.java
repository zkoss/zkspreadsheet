/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/


//right click number format of row 12
public class SS_118_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		verifyFalse(jq("$_formatNumberDialog").isVisible());
		
		rightClickRowHeader(11);
		click(jq("$numberFormat a.z-menu-item-cnt"));
		waitResponse();

		//verify
		verifyTrue(jq("$_formatNumberDialog").isVisible());
	}
}



