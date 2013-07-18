/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//select B4:H10, and CTRL+B to make font bold
public class SS_163_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(1,3,7,9);

		//B:66
		pressCtrlWithChar("66");
		
		//verify
		String style = getCellStyle(1, 6);
		verifyTrue(containsIgnoreCase(style, "font-weight: 700"));		
		style = getCellStyle(1, 7);
		verifyTrue(containsIgnoreCase(style, "font-weight: 700"));
		style = getCellStyle(5, 6);
		verifyTrue(containsIgnoreCase(style, "font-weight: 700"));
	}
}

