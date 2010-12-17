/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//right click "Border" to have popup menu: B13
public class SS_150_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCells(1,10,5,12);

		//fail to trigger???
//		clickAt(jq(".z-dpbutton-arrow:eq(4)"),"5,5");
//		waitResponse();

//		click(jq(".z-dpbutton-arrow:eq(4)"));
//		waitResponse();
		
		mouseDownAt(jq(".z-dpbutton-arrow:eq(4)"),"5,5");
		waitResponse();
		mouseUpAt(jq(".z-dpbutton-arrow:eq(4)"),"5,5");
		waitResponse();
		//how to verify
		sleep(5000);
	}
}

