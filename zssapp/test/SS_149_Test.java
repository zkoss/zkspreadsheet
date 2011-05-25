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

//right click "fill color" : B13
public class SS_149_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		rightClickCell(1,12);
		click(jq("div[title=\"Fill color\"] img.z-colorbtn-btn:eq(2)"));
		waitResponse();

	    // Input color hex code, then press Enter.
        JQuery colorTextbox = jq(".z-colorbtn-pp:visible .z-colorpalette-hex-inp");
        String bgColorStr = "#00ff00";
        type(colorTextbox, bgColorStr);
        keyPressEnter(colorTextbox);
        
        //Verify
        JQuery cell_B_13_Outer = getSpecifiedCellOuter(1, 12);
        String style = cell_B_13_Outer.css("background-color");

        //input "#00ff00", but it actually get "009900"
        //Is it acceptable in this spec?
        if (style != null) {
            verifyTrue("Unexcepted result: " + cell_B_13_Outer.css("background-color"), 
            		ColorVerifingHelper.isEqualColor("#00ff00", style));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
	}
}

