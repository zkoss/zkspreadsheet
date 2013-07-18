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
	
	/**
	 * Auto fill
	 */
	@Override
	protected void executeTest() {

		String J12_Text = getCellText(9, 11);
		//TODO: decode hex color format to RGB
		//String J12_Background_Color = getCellBackgroundColor(9, 11);
		
		String K12_Text = getCellText(10, 11);
		verifyTrue(J12_Text.length() > 0 && "".equals(K12_Text));
		
		String K12_Background_Color = getCellBackgroundColor(10, 11);
		verifyTrue("rgb(255, 255, 255)".equals(K12_Background_Color));
		
		dragFill(9, 11, 11, 11);
		/**
		 * Expect:
		 * 
		 * Copy the cell value and style
		 */
		String filled_K12 = getCellText(10, 11);
		verifyTrue("drag fill shall copy the cell text from source", J12_Text.equals(filled_K12));
		String filled_K12_Background_Color = getCellBackgroundColor(10, 11);
		verifyTrue("drag fill shall copy the cell style from source", "rgb(141, 179, 226)".equals(filled_K12_Background_Color));
	}
}



