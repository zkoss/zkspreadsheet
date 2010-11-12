/* Zssapps.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 12, 2010 11:35:45 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul;

import java.util.HashMap;

import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author Sam
 *
 */
public final class ZssappComponents {
	private ZssappComponents(){};
	
	private final static String KEY_ARG_SPREADSGEET = "org.zkoss.zss.app.zssappComponent.spreadsheet";
	
	public static HashMap newSpreadsheetArg(Spreadsheet spreadsheet) {
		HashMap<String, Object> arg = new HashMap<String, Object>();
		arg.put(KEY_ARG_SPREADSGEET, spreadsheet);
		return arg;
	}
	
	
	public static Spreadsheet getSpreadsheetFromArg() {
		return (Spreadsheet)Executions.getCurrent().getArg().get(KEY_ARG_SPREADSGEET);
	}
}
