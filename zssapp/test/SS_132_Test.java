/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//delete => entire column : F12
public class SS_132_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		//verify
		String oriF12value = getSpecifiedCell(5,11).text();
		String oriG12value = getSpecifiedCell(6,11).text();		
		verifyNotEquals(oriF12value, oriG12value);
		String oriF13value = getSpecifiedCell(5,12).text();
		String oriG13value = getSpecifiedCell(6,12).text();		
		verifyNotEquals(oriF13value, oriG13value);
		
		
		rightClickCell(5,11);
		mouseOver(jq("a.z-menu-cnt:eq(1)"));		
		waitResponse();
		click(jq("$deleteEntireColumn a.z-menu-item-cnt"));
		waitResponse();
		
		//verify
		String f12value = getSpecifiedCell(5,11).text();
		verifyEquals(f12value, oriG12value);
		String f13value = getSpecifiedCell(5,12).text();
		verifyEquals(f13value, oriG13value);
		
	}
}



