/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//select B4:H10, and CTRL+Delete to clear style
public class SS_167_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(1,3,7,9);

		pressCtrlWithChar(DELETE_NATIVE);
		
		//verify
		String style = getCellCompositeStyle(1, 6);
		verifyEquals(style,"rgba(0, 0, 0, 0):Arial:left");		
		style = getCellCompositeStyle(1, 7);
		verifyEquals(style,"rgba(0, 0, 0, 0):Arial:left");
		style = getCellCompositeStyle(5, 6);
		verifyEquals(style,"rgba(0, 0, 0, 0):Arial:right");
		sleep(5000);
	}
}

