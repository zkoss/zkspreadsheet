/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//cell K1, find function with "bin"
public class SS_204_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {		
		rightClickCell(10,0);
		
		click(jq("$formula a.z-menu-item-cnt"));
		waitResponse();
		type(jq("$searchTextbox"), "bin");
		waitResponse();
		click(jq("$searchBtn td.z-button-cm"));
		waitResponse();

		//how to verify
		sleep(5000);
	}
}



