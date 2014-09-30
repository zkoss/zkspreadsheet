/* ZSSBy.java

	Purpose:
		
	Description:
		
	History:
		Mon, Jul 21, 2014  6:18:00 PM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.test.selenium.entity;

import org.openqa.selenium.By;

/**
 * 
 * @author RaymondChao
 */
public class ZSSBy {
	/**
	 * Returns a {@code By} which locates elements by the JavaScript expression passed to it.
	 * 
	 * @param script The JavaScript expression to run and whose result to return
	 */
	public static By javascript(String script) {
		return new ByJavaScript(script);
	}
}
