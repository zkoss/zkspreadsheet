/* SpreadsheetLabel.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 1, 2012 9:40:31 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui;

import java.util.ArrayList;
import java.util.List;

/**
 * Label's key of spreadsheet
 * 
 * @author sam
 *
 */
public class SpreadsheetLabel {
	private SpreadsheetLabel() {};
	
	public enum Sheet {
		SHEET("sheet"),
		ADD("add"),
		DELETE("delete"),
		RENAME("rename"),
		MOVE_LEFT("moveLeft"),
		MOVE_RIGHT("moveRight"),
		PROTECT("protect");
		
		private final String key;
		private Sheet(String key) {
			this.key = key;
		}
		
		public String getLabelKey() {
			return "zss.sheet." + key;
		}
		
		@Override
		public String toString() {
			return key;
		}
	}
	
	private static ArrayList<String> labelKeys;
	static {
		labelKeys = new ArrayList<String>();
		
		for (Sheet key : Sheet.class.getEnumConstants()) {
			labelKeys.add(key.getLabelKey());
		}
	}

	public static List<String> getLabelKeys() {
		return labelKeys;
	}
}