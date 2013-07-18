/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//select B13:G18, and Delete to clear content
public class SS_166_Test extends SSAbstractTestCase {
	
	/**
	 * "Delete" key: clear cell style and cell text
	 */
	@Override
	protected void executeTest() {
		selectCells(1,12,6,17);

		//Delete is confused with ".", should use native
		keyPressNative(DELETE_NATIVE);
		waitResponse();
		keyUpNative(DELETE_NATIVE);
		waitResponse();
		
		//verify
		String b13value = getSpecifiedCell(1,12).text();
		verifyTrue("".equals(b13value));

		String g18value = getSpecifiedCell(6,17).text();
		verifyTrue("".equals(g18value));
	}
}

