import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;

/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

public class SS_148_Test extends SSAbstractTestCase {
	
	/**
	 * Set font color (from fast toolbar)
	 */
	@Override
	protected void executeTest() {
        String selectedColor = setCellFontColorByFastToolbarbutton(1, 12, 98);
        String newFontColor = getCellFontColor(1, 12);
        
        //TODO: decode # format and compare color
        verifyEquals("rgb(153, 102, 255)", newFontColor);
	}
}

