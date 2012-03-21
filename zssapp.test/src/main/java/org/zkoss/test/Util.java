/* Util.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 13, 2012 2:07:34 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import java.lang.reflect.Method;

/**
 * @author sam
 *
 */
public class Util {
	private Util() {};
	
	public static Integer intValue(Object number) {
		if (number instanceof Number) {
			if (number instanceof Integer) {
				return (Integer)number;
			} else {
				Method method = null;
				try {
					method = number.getClass().getMethod("intValue");
					return (Integer)method.invoke(number);
				} catch (Exception ex) {
					throw new RuntimeException("can't find intValue method");
				}
			}
		}
		return null;
	}
}
