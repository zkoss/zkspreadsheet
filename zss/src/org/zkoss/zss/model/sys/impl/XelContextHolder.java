/* XelContextHolder.java

	Purpose:
		
	Description:
		
	History:
		Apr 7, 2010 9:46:43 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.sys.impl;

import org.zkoss.xel.XelContext;

/**
 * Helper class to pass XelContext for formula evaluation. 
 * @author henrichen
 * @see BookHelper#evaluate
 */
public class XelContextHolder {
	private static ThreadLocal<XelContext> _ctx = new ThreadLocal<XelContext>();
	
	public static void setXelContext(XelContext value) {
		_ctx.set(value);
	}
	
	public static XelContext getXelContext() {
		return _ctx.get();
	}
}
