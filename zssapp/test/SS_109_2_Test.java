/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/


//hide column F
public class SS_109_2_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		//verify
		verifyTrue(getColumnHeader(5).isVisible());
		
		rightClickColumnHeader(5);
		click(jq("$hide a.z-menu-item-cnt"));
		waitResponse();

		//verify
		verifyFalse(getColumnHeader(5).isVisible());		
	}
}



