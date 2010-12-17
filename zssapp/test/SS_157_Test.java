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
public class SS_157_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		//verify
		verifyTrue(jq("div.zshighlight").width() == 0);

		selectCells(5,5,8,8);		
		pressCtrlWithChar(C);
		
		//verify
		verifyTrue(jq("div.zshighlight").width() != 0);
		
		selectCells(10, 5, 10, 5);
		pressCtrlWithChar(V);
		
		//verify
		String f6value = getSpecifiedCell(5,5).text();
		String f7value = getSpecifiedCell(5,6).text();
		
		String k6value = getSpecifiedCell(10,5).text();
		String k7value = getSpecifiedCell(10,6).text();
		
		verifyEquals(f6value, k6value);
		verifyEquals(f7value, k7value);
	}
}

