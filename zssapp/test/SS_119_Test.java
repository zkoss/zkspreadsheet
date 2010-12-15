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
import org.zkoss.ztl.ZKClientTestCase;

import com.thoughtworks.selenium.Selenium;


public class SS_119_Test extends ZKClientTestCase {
	
	public SS_119_Test() {
		target = "http://zktest/zssdemos/index.zul";
		browsers = getBrowsers("chrome");
//		browsers = getBrowsers("firefox");
		_timeout = 60000;
	}
		
	@Test(expected = AssertionError.class)
	public void testClick() {
		for (Selenium browser : browsers) {
			try {
				start(browser);
				windowFocus();
				windowMaximize();
				
				mouseDown(jq("div.zsleftcell[z\\\\.r=\"11\"] div"));
				waitResponse();
				mouseUp(jq("div.zsleftcell[z\\\\.r=\"11\"] div"));
				waitResponse();
				mouseDown(jq("div.zsleftcell[z\\\\.r=\"11\"] div"));
				waitResponse();
				mouseUp(jq("div.zsleftcell[z\\\\.r=\"11\"] div"));
				waitResponse();
				contextMenuAt(jq("div.zsleftcell[z\\\\.r=\"11\"] div"),"2,2");
				waitResponse();
				click(jq("$hide a.z-menu-item-cnt"));
				waitResponse();
				//how to verify
				sleep(5000);
			} finally {
				stop();	
			}
		}
	}
}



