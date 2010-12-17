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

		pressCtrlWithChar(DELETE);
		
		//verify
		sleep(5000);
//		String style = getCellStyle(1, 6);
//		verifyTrue(style.contains("text-decoration: underline;"));		
//		style = getCellStyle(1, 7);
//		verifyTrue(style.contains("text-decoration: underline;"));
//		style = getCellStyle(5, 6);
//		verifyTrue(style.contains("text-decoration: underline;"));
	}
}

