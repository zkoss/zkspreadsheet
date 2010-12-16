/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//right click "Number Format" of column F
public class SS_109_1_Test extends SSAbstractTestCase {
	
	@Override
	protected void executeTest() {
		rightClickColumnHeader(5);
		click(jq("$numberFormat a.z-menu-item-cnt"));
		waitResponse();

		//verify
		String titleOfPopup =  jq(".z-window-highlighted.z-window-highlighted-shadow .z-window-highlighted-header").attr("textContent");
		verifyEquals(titleOfPopup,"Number Format");		
	}
}



