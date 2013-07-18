/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//select B4:H10, and CTRL+I to make font italic
public class SS_164_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(1,3,7,9);

		//I:73		
		pressCtrlWithChar("73");
		
		//verify
		String style = getCellStyle(1, 6);
		verifyTrue(containsIgnoreCase(style, "font-style: italic"));		
		style = getCellStyle(1, 7);
		verifyTrue(containsIgnoreCase(style, "font-style: italic"));
		style = getCellStyle(5, 6);
		verifyTrue(containsIgnoreCase(style, "font-style: italic"));
	}
}

