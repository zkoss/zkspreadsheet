/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//cut cell F12
public class SS_121_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		//verify
		verifyTrue(jq("div.zshighlight") == null);
		
		rightClickCell(5,11);
		click(jq("$cut a.z-menu-item-cnt"));
		waitResponse();

		//verify
		verifyTrue(jq("div.zshighlight") != null);
	}
}



