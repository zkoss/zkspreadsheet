/* ZssVariableResolver.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 1, 2008 7:11:37 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.demo.test;

import org.zkoss.xel.VariableResolver;
import org.zkoss.xel.XelException;
import org.zkoss.zss.demo.ReportContext;
import org.zkoss.zss.demo.UserBean;

/**
 * @author Dennis.Chen
 *
 */
public class ZssVariableResolver implements VariableResolver{

	UserBean userbean = new UserBean("Dennis","Chen");
	ReportContext report = new ReportContext();

	public Object resolveVariable(String name) throws XelException {
		if(name.equals("userbean")){
			return userbean;
		}else if(name.equals("reportContext")){
			return report;
		}
		return null;
	}

}
