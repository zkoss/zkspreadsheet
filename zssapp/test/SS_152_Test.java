/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//right click "center text" : B13
public class SS_152_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCell(1,12);
		click(jq(".z-toolbarbutton[title=\"Center Text\"] img:eq(2)"));
		waitResponse();
		rightClickCell(1,13);

		//verify
		String style = getCellStyle(1, 12);
		verifyTrue(containsIgnoreCase(style, "text-align: center"));		
	}
}

