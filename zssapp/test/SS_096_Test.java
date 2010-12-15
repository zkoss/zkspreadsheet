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


public class SS_096_Test extends SSAbstractTestCase {
	
	@Override
	protected void executeTest() {
		contextMenu(jq("@tab[label=\"Market\"] span.z-tab-text"));				
		waitResponse();
		click(jq("$shiftSheetLeft a.z-menu-item-cnt"));
		waitResponse();

		//how to verify?
		sleep(5000);
//		Widget marketSibling = widget(jq("@tab[label=\"Market\"] span.z-tab-text")).nextSibling();
//		waitResponse();
//		Widget debt = widget(jq("@tab[label=\"Debt\"] span.z-tab-text"));
//		waitResponse();
		
	}
}



