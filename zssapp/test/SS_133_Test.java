/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//clear -> clear content : F12
public class SS_133_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCell(5,11);
		mouseOver(jq("a.z-menu-cnt:eq(2)"));		
		waitResponse();
		click(jq("$clearContent a.z-menu-item-cnt"));
		waitResponse();
		
		//how to verify
		sleep(5000);
	}
}



