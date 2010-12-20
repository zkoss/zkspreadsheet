/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//select F6:I9, and CTRL+D to clear content
public class SS_158_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		//verify
		String beforeF6Style = getCellCompositeStyle(5, 5);

		selectCells(5,5,8,8);

		//D:68
		pressCtrlWithChar("68");
		
		//verify
		String afterF6Style = getCellCompositeStyle(5, 5);
		String f6value = getSpecifiedCell(5,5).text();
		verifyEquals(f6value,null);

		String f7value = getSpecifiedCell(5,6).text();
		verifyEquals(f7value,null);
		
		String g8value = getSpecifiedCell(6,7).text();
		verifyEquals(g8value,null);
		
		verifyEquals(beforeF6Style,afterF6Style);

	}
}

