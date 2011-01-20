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

//right click to shift right sheet "Debt"
public class SS_097_Test extends SSAbstractTestCase {
			
	@Override
	protected void executeTest() {
		contextMenu(jq("@tab[label=\"Debt\"] span.z-tab-text"));				
		waitResponse();
		click(jq("$shiftSheetRight a.z-menu-item-cnt"));
		waitResponse();

		//verify
		Widget marketSibling = widget(jq("@tab[label=\"Market\"]")).nextSibling();
		String content = jq(marketSibling).find("span").text();		
		verifyEquals("Debt", content);		
	}
}



