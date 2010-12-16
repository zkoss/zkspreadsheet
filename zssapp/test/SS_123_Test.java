/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//copy cell F12 to G12
public class SS_123_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		//verify
		String f12value = getSpecifiedCell(5,11).text();
		String g12value = getSpecifiedCell(6,11).text();
		verifyNotEquals(f12value, g12value);
		verifyTrue(jq("div.zshighlight").width() == 0);
		
		
		rightClickCell(5,11);
		click(jq("$copy a.z-menu-item-cnt"));
		waitResponse();

		//verify
		verifyTrue(jq("div.zshighlight").width() != 0);

		rightClickCell(6,11);
		click(jq("$paste a.z-menu-item-cnt"));
		waitResponse();

		//verify
		f12value = getSpecifiedCell(5,11).text();
		g12value = getSpecifiedCell(6,11).text();
		verifyEquals(f12value, g12value);
	}
}



