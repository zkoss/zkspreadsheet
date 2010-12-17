/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//select B4:H10, and Delete to clear content
public class SS_166_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(1,3,7,9);

		//not work, why?
		keyPressNative(DELETE);
		waitResponse();
		keyUpNative(DELETE);
		waitResponse();
		
		//verify
		sleep(5000);
	}
}

