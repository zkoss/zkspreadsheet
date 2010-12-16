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
		int width = getColumnHeader(5).width();
		verifyTrue(width!=0);
		
		rightClickColumnHeader(5);
		click(jq("$hide a.z-menu-item-cnt"));
		waitResponse();

		//verify
		width = getColumnHeader(5).width();
		verifyTrue(width==0);		
	}
}



