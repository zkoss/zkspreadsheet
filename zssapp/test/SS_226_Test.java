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
		//clickDropdownButtonMenu("$fastIconBtn $sortDropdownBtn", 3);
		clickDropdownButtonMenu("$fastIconBtn $sortDropdownBtn", "Filter");
		int btmSize = getRow(5).children(".zsbtn").length();
		verifyEquals(btmSize, 6);

		//TODO: verify buttons position
		//clickCell(getSpecifiedCell(2, 6));
		
		//toggle autofilter off
		clickCell(F11);
		clickDropdownButtonMenu("$fastIconBtn $sortDropdownBtn", "Filter");
		verifyEquals(getRow(5).children(".zsbtn").length(), 0);
	}

}
