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

public class SS_150_Test extends SSAbstractTestCase {
	
	/**
	 * Open border menu (from fast toolbar)
	 */
	@Override
	protected void executeTest() {
		rightClickCells(1,10,5,12);

        // Click Border icon
        JQuery borderIcon = jq("$borderBtn:eq(2)");
        mouseOver(borderIcon);
        waitResponse();
        
        //TODO: didn't find the appropriate offset for IE8 to click
        //the same code works in case 69, 
        //it's very similar to case 69, why not work here?
        clickAt(borderIcon, "30,0");
        waitResponse();
		
		
        /**
         * Expected
         * 
         * Open set border menu
         */
        String topBorder = jq(".z-menu-popup:visible .z-menu-item:eq(1)").text();
        verifyEquals(topBorder, " Top border");
	}
}












