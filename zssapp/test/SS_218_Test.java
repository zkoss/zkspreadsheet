/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//enter "12345.6789" in cell H1, and set number format to "Scientific","scientific"
public class SS_218_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(7, 0, 7, 0);
		type(jq("$formulaEditor"), "12345.6789");
		waitResponse();
	
		rightClickCells(7,0,7,0);
		waitResponse();
		click(jq("$format a.z-menu-item-cnt"));
		waitResponse();
		
		click(jq("@listcell[label=\"Scientific\"] div.z-overflow-hidden"));
		waitResponse();
		click(jq("@listcell[label=\"scientific\"] div.z-overflow-hidden"));
		waitResponse();
		
		click(jq("$okBtn img"));
		waitResponse();
		
		//verify
		String h1value = getSpecifiedCell(7,0).text();
		verifyEquals(h1value,"1.23E+04");
	}
}



