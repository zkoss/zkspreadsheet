/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/


//clear style of row 12
public class SS_114_Test extends SSAbstractTestCase {

	@Override
	protected void executeTest() {
		//verify
		String beforeF12Style = getCellCompositeStyle(5, 11);
		String beforef12value = getSpecifiedCell(5,11).text();
		
		
		rightClickRowHeader(11);
		click(jq("$clearStyle a.z-menu-item-cnt"));
		waitResponse();

		//verify
		String afterF12Style = getCellCompositeStyle(5, 11);
		String afterf12value = getSpecifiedCell(5,11).text();
		
		verifyEquals(beforef12value,afterf12value);
		verifyEquals(afterF12Style, CELL_WITHOUT_STYLE);
		
		sleep(5000);		
	}
}



