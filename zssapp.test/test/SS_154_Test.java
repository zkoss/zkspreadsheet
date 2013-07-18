/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//right click "merge" : F12:H12
public class SS_154_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		//verify
		int beforeF12width = getSpecifiedCell(5, 11).width();
		int beforeG12width = getSpecifiedCell(6, 11).width();
		int beforeH12width = getSpecifiedCell(7, 11).width();
		
		String f12Text = getSpecifiedCell(5, 11).text();
		
		rightClickCells(5,11,7,11);
		click(jq("$_mergeCellBtn img"));
		waitResponse();
		
		//verify
		int afterF12width = getSpecifiedCell(5, 11).width();
		
		String mergeText = getSpecifiedCell(5, 11).text();
		
		verifyTrue(afterF12width > beforeF12width+beforeG12width+beforeH12width);
		verifyFalse(getSpecifiedCell(6, 11).isVisible());
		verifyFalse(getSpecifiedCell(7, 11).isVisible());
		verifyEquals(f12Text, mergeText);
	}
}

