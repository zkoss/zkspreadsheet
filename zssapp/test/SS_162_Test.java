/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//CTRL + O to open new file
public class SS_162_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		verifyFalse(jq("$_openFileDialog").isVisible());
		
		selectCells(5,5,5,5);

		//O:79
		pressCtrlWithChar("79");
		
		//verify
		verifyTrue(jq("$_openFileDialog").isVisible());
		
		//TODO: in IE, will open browser dialog
	}
}

