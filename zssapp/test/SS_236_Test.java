import org.zkoss.ztl.JQuery;

/* SS_236_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		May 26, 2011 4:16:22 PM , Created by sam
}}IS_NOTE

Copyright (C) 2011 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
 */

/**
 * @author sam
 *
 */
public class SS_236_Test extends SSAbstractTestCase {		
	
	/**
	 * Steps
	 * 1. without sheet protection, keyin "1"
	 * 2. click protect sheet button
	 * 3. keyin "2", can not edit cell value. value remain "1"
	 */
	@Override
	protected void executeTest() {
		
		JQuery A0 = getSpecifiedCell(0, 0);
		clickCell(A0);
		clickCell(A0);
		typeKeys(A0, "1");
		keyPressEnter(jq(".zsedit"));
		
		click(jq("$fastIconBtn $protectSheet input"));
		waitResponse();
		
		clickCell(A0);
		typeKeys(A0, "2");
		keyPressEnter(jq(".zsedit"));
		waitResponse();
		verifyEquals("1", A0.text());
	}

}
