/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/


//clear style of column F
public class SS_105_Test extends SSAbstractTestCase {
	
	@Override
	protected void executeTest() {
		String beforef12value = getSpecifiedCell(5,11).text();
		rightClickColumnHeader(5);
		click(jq("$clearStyle a.z-menu-item-cnt"));
		waitResponse(20000);

		//how to verify		
		String afterf12value = getSpecifiedCell(5,11).text();
		verifyEquals(beforef12value,afterf12value);

		String styleAfterF12 = getCellStyle(5,11);
		System.out.println(">>>style:"+styleAfterF12);
	}
}



