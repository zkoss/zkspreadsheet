/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//use "shift" to select an area
public class SS_221_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(1, 0, 1, 0);
		shiftKeyDown();
		//RIGHT in awt event is 1007
		keyDownNative(A);
		keyUpNative(A);
		shiftKeyUp();
		
		//how to verify
	}
}



