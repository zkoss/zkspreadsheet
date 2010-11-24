/* Zssapps.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 12, 2010 11:35:45 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.HashMap;

import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author Sam
 *
 */
public class Zssapps {
	private Zssapps(){};
	
	private final static String KEY_ARG_SPREADSGEET = "org.zkoss.zss.app.zssappComponent.spreadsheet";
	
	//TODO: use constructor to pass spreadsheet to dialog, remove this
	public static HashMap newSpreadsheetArg(Spreadsheet spreadsheet) {
		HashMap<String, Object> arg = new HashMap<String, Object>();
		arg.put(KEY_ARG_SPREADSGEET, spreadsheet);
		return arg;
	}
	
	public static Spreadsheet getSpreadsheetFromArg() {
		return (Spreadsheet)Executions.getCurrent().getArg().get(KEY_ARG_SPREADSGEET);
	}
	
	//TODO: remove this mechanism
	@Deprecated
	public static void bindSpreadsheet(Spreadsheet spreadsheet, Object target) {
		Field[] flds = target.getClass().getDeclaredFields();
		for (Field f : flds) {
			final boolean old = f.isAccessible();
			try {
				f.setAccessible(true);
				Object obj = f.get(target);
				if (obj instanceof ZssappComponent) {
					Method m = f.getType().getMethod("bindSpreadsheet", Spreadsheet.class);
					m.invoke(obj, spreadsheet);
				}
			} catch (Exception ex) {
				ex.printStackTrace();
			} finally {
				f.setAccessible(old);
			}
		}
	}
}
