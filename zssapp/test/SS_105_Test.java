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
		//verify
		String beforeF12Style = getCellCompositeStyle(5, 11);
		String beforef12value = getCellText(5,11);
		
		rightClickColumnHeader(5);
		click(jq("$clearStyle a.z-menu-item-cnt"));
		waitResponse(20000);

		//verify
		String afterF12Style = getCellCompositeStyle(5, 11);
		String afterf12value = getCellText(5,11);
		
		verifyEquals(beforef12value, afterf12value);
		verifyTrue(CELL_WITHOUT_STYLE.equals(afterF12Style) || CELL_WITHOUT_STYLE2.equals(afterF12Style));
	}
}



