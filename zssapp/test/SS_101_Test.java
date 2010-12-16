/* order_test_1Test.java

	Purpose:
		
	Description:
		
	History:
		Sep, 7, 2010 17:30:59 PM

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

This program is distributed under Apache License Version 2.0 in the hope that
it will be useful, but WITHOUT ANY WARRANTY.
*/

import org.junit.Test;
import org.zkoss.ztl.Element;
import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.Tags;
import org.zkoss.ztl.Widget;
import org.zkoss.ztl.ZK;
import org.zkoss.ztl.ZKClientTestCase;

import com.thoughtworks.selenium.Selenium;


//cut column F
public class SS_101_Test extends SSAbstractTestCase {
	
	@Override
	protected void executeTest() {
		rightClickColumnHeader(5);
		click(jq("$cut a.z-menu-item-cnt"));
		waitResponse();

		//verify
		verifyTrue(jq("div.zshighlight") != null);		
	}
}



