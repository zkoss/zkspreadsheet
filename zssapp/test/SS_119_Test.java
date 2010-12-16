/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/


//hide row 12
public class SS_119_Test extends SSAbstractTestCase {
	
	@Override
	protected void executeTest() {
		//verify
		int height = getRowHeader(11).height();
		verifyTrue(height!=0);
		
		rightClickRowHeader(11);
		click(jq("$hide a.z-menu-item-cnt"));
		waitResponse();
		//verify
		height = getRowHeader(11).height();
		verifyTrue(height==0);		
	}
}



