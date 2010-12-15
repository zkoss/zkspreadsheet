/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//copy from Column F to Column G
public class SS_103_Test extends SSAbstractTestCase {	

	@Override
	protected void executeTest() {
		rightClickColumnHeader(5);
		click(jq("$copy a.z-menu-item-cnt"));
		waitResponse();
		rightClickColumnHeader(6);
		click(jq("$paste a.z-menu-item-cnt"));
		waitResponse();
		//how to verify
		sleep(5000);
	}
}



