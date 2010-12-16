/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//clear content of column F
public class SS_104_Test extends SSAbstractTestCase {	
	@Override
	protected void executeTest() {
		rightClickColumnHeader(5);
		String styleBeforeF12 = getCellStyle(5,11);
		click(jq("$clearContent a.z-menu-item-cnt"));
		waitResponse();
		//how to verify
		String f9value = getSpecifiedCell(5,8).text();
		verifyEquals(f9value,null);
		
		String styleAfterF12 = getCellStyle(5,11);
		verifyEquals(styleBeforeF12,styleAfterF12);
	}
}



