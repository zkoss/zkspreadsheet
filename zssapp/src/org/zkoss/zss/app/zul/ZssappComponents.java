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
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.EventQueues;
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
	
	public static void publishFontSizeChanged(Spreadsheet spreadsheet, String fontSize) {
		
		EventQueues.
			lookup(spreadsheet.getId() + "$FontSizeChanged", EventQueues.DESKTOP, true).
			publish(new Event("onFontSizeChanged", null, fontSize));
	}
	
	public static void subscribeFontSizeChanged(Spreadsheet spreadsheet, EventListener listener) {
		EventQueues.
			lookup(spreadsheet.getId() + "$FontSizeChanged", EventQueues.DESKTOP, true).subscribe(listener);
	}
	
	public static void publishFontFamilyChanged(Spreadsheet spreadsheet, String fontFamily) {
		EventQueues.
			lookup(spreadsheet.getId() + "$FontFamilyChanged",  EventQueues.DESKTOP, true).
			publish(new Event("onFontFamilyChanged", null, fontFamily));
	}
	
	public static void subscribeFontFamilyChanged(Spreadsheet spreadsheet, EventListener listener) {
		EventQueues.
			lookup(spreadsheet.getId() + "$FontFamilyChanged", EventQueues.DESKTOP, true).subscribe(listener);
	}
	
	public static void publishFontBoldChanged(Spreadsheet spreadsheet, Boolean bold) {
		EventQueues.
			lookup(spreadsheet.getId() + "$FontBoldChanged",  EventQueues.DESKTOP, true).
			publish(new Event("onFontBoldChanged", null, bold));
	}
	
	public static void subscribeFontBoldChanged(Spreadsheet spreadsheet, EventListener listener) {
		EventQueues.
			lookup(spreadsheet.getId() + "$FontBoldChanged", EventQueues.DESKTOP, true).subscribe(listener);
	}
	
//	public static void bindSpreadsheetController(Spreadsheet spreadsheet, Object controller) {
//		
//	}
//	
//	public static Object getSpreadsheetController(Spreadsheet spreadsheet) {
//		return 
//	}
}