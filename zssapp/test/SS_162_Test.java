/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//select F6:I9, and CTRL+C, and copy to K6
public class SS_162_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(5,5,5,5);

		//O:79
		pressCtrlWithChar("79");
		
		//verify
		String titleOfPopup =  jq(".z-window-highlighted.z-window-highlighted-shadow .z-window-highlighted-header").attr("textContent");
		verifyEquals(titleOfPopup,"Open File");		

	}
}

