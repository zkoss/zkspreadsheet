import org.zkoss.ztl.JQuery;

/* SS_226_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		May 26, 2011 6:31:41 PM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */

/**
 * @author sam
 *
 */
public class SS_226_Test extends SSAbstractTestCase {

	/**
	 * Steps
	 * 1. toggle on autofilter
	 * 2. toggle off autofilter
	 */
	@Override
	protected void executeTest() {
		JQuery F11 = getSpecifiedCell(5, 11);
		clickCell(F11);
		clickCell(F11);
		
		//click filter menu
		clickDropdownButtonMenu("$fastIconBtn $sortDropdownBtn div.z-dpbutton-btn", "Filter");
		waitResponse();
		
		int btmSize = getRow(11).children(".zsbtn").length();
		verifyEquals(6, btmSize);

		//TODO: verify buttons position
		//clickCell(getSpecifiedCell(2, 6));
		
		//toggle auto filter off
		clickCell(F11);
		clickDropdownButtonMenu("$fastIconBtn $sortDropdownBtn div.z-dpbutton-btn", "Filter");
		verifyEquals(getRow(11).children(".zsbtn").length(), 0);
	}

}
