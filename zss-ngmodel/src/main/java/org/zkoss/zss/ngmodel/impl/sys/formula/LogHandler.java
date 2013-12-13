/* LogHandler.java

	Purpose:
		
	Description:
		
	History:
		Nov 19, 2013 Created by Pao Wang

Copyright (C) 2013 Potix Corporation. All Rights Reserved.
 */
package org.zkoss.zss.ngmodel.impl.sys.formula;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;

/**
 * @author Pao
 */
public class LogHandler implements InvocationHandler {
	public static boolean debug = false;
	private Object delegatee;

	public LogHandler(Object delegatee) {
		this.delegatee = delegatee;
	}

	@Override
	public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
		if(debug) {
			StringBuilder sb = new StringBuilder();
			sb.append(delegatee.getClass().getSimpleName()).append('.');
			sb.append(method.getName()).append('(');
			for(Class<?> p : method.getParameterTypes()) {
				sb.append(p.getSimpleName()).append(", ");
			}
			if(sb.charAt(sb.length() - 1) == '(') {
				sb.append(')');
			} else {
				sb.delete(sb.length() - 2, sb.length()).append(')');
			}
			System.out.println(sb);
		}
		return method.invoke(delegatee, args);
	}

}
