/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/


//copy row 12
public class SS_111_Test extends SSAbstractTestCase {
	
	@Override
	protected void executeTest() {
		rightClickRowHeader(11);
		click(jq("$copy a.z-menu-item-cnt"));
		waitResponse();

		//verify
		verifyTrue(jq("div.zshighlight") != null);		
	}
}



