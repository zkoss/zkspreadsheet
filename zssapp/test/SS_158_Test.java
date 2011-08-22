/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

public class SS_158_Test extends SSAbstractTestCase {
	
	
	/**
	 * use Ctrl + D to clear cell
	 */
	@Override
	protected void executeTest() {
		//verify
		String beforeF6Style = getCellCompositeStyle(5, 5);

		selectCells(5,5,8,8);

		//D:68
		pressCtrlWithChar("68");
		
		//verify
		String afterF6Style = getCellCompositeStyle(5, 5);
		String f6value = getCellText(5,5);

		String f7value = getCellText(5,6);
		verifyTrue("Clear cell: the cell text shall be empty", "".equals(f7value));
		
		String g8value = getCellText(5,7);
		verifyTrue("Clear cell: the cell text shall be empty", "".equals(g8value));
		
		verifyEquals(beforeF6Style,afterF6Style);

	}
}

