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


public class SSF6Delete extends ZKClientTestCase {
	
	public SSF6Delete() {
		target = "http://zktest/zssdemos/index.zul";
		browsers = getBrowsers("chrome");
//		browsers = getBrowsers("firefox");.
		_timeout = 60000;
	}
		
	@Test(expected = AssertionError.class)
	public void testClick() {
		for (Selenium browser : browsers) {
			try {
				start(browser);
				windowFocus();
				windowMaximize();
				
				mouseDownAt(jq("div.zscell[z\\\\.r=\"5\"][z\\\\.c=\"5\"] div"),"5,5");				 
				waitResponse();
				mouseUpAt(jq("div.zscell[z\\\\.r=\"5\"][z\\\\.c=\"5\"] div"),"5,5");
				waitResponse();
				mouseDownAt(jq("div.zscell[z\\\\.r=\"5\"][z\\\\.c=\"5\"] div"),"5,5");
				waitResponse();
				mouseUpAt(jq("div.zscell[z\\\\.r=\"5\"][z\\\\.c=\"5\"] div"),"5,5");
				waitResponse();
				
				keyPressNative(DELETE);
				waitResponse();
				keyUpNative(DELETE);
				waitResponse();
				sleep(5000);
			} finally {
				stop();	
			}
		}
	}
}



