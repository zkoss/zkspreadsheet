/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//delete column F
public class SS_107_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickColumnHeader(5);
		click(jq("$deleteColumn a.z-menu-item-cnt"));
		waitResponse();
		
		//verify
		String f12value = getSpecifiedCell(5,11).text();
		verifyEquals(f12value,"Q2");
	}
}