/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/


//insert column before Column F
public class SS_106_Test extends SSAbstractTestCase {
	
	@Override
	protected void executeTest() {
		rightClickColumnHeader(5);
		click(jq("$insertColumn a.z-menu-item-cnt"));
		waitResponse();
		
		//verify
		String f12value = getSpecifiedCell(5,11).text();
		String g12value = getSpecifiedCell(6,11).text();
		verifyEquals(f12value,null);
		verifyEquals(g12value,"Q1");
	}
}



