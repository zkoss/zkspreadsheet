import org.zkoss.ztl.JQuery;

/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

//use mouse to auto fill
public class SS_222_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(9,11,9,11);
		JQuery cellJ12 = getSpecifiedCell(9, 11);
		int width = cellJ12.width();
		int height = cellJ12.height();
		mouseOver(jq(".zsselecti"));
		waitResponse();
		mouseDownAt(jq(".zsselecti"), ""+width+","+height);
		waitResponse();
		mouseMoveAt(getSpecifiedCell(11, 11), "2,2");
		waitResponse();
		mouseUpAt(getSpecifiedCell(11, 11), "2,2");
		waitResponse();
		
		//how to verify
	}
}



