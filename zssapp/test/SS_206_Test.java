/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//cell K1, choose function "ACCRINT"
public class SS_206_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {		
		rightClickCell(10,0);
		
		click(jq("$formula a.z-menu-item-cnt"));
		waitResponse();
		click(jq("@listcell[label=\"ACCRINT\"]"));		
		waitResponse();
		
		//verify
		String funcDesc = jq(".z-vlayout-inner .z-label:eq(1)").text();
		String expect = "ACCRINT(issue, first_interest, settlement, rate, par, frequency, basis, calc_method)";
		verifyEquals(funcDesc, expect);
	}
}



