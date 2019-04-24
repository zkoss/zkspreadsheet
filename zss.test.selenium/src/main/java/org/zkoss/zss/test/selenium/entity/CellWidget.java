/* CellWidget.java

	Purpose:
		
	Description:
		
	History:
		Mon, Jul 21, 2014 11:37:03 AM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.test.selenium.entity;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 
 * @author RaymondChao
 */
public class CellWidget extends Widget {
	private static final String CELL = ".getCell(%1$d,%2$d)";
	private static final Pattern refPattern = 
			Pattern.compile("([A-Z]+)(\\d+)");
	public CellWidget(SheetCtrlWidget spreadsheet, int row, int column) {
		setSelector(spreadsheet.getSelector() + String.format(CELL, row, column));
	}
	
	public CellWidget(SheetCtrlWidget spreadsheet, String cellRef) {
		if (isEmpty(cellRef)) {
			throw new NullPointerException("cell reference cannot be null.");
		}
		final Matcher refMatcher = refPattern.matcher(cellRef);
		if (refMatcher.matches()) {
			final int column = toColumn(refMatcher.group(1)) - 1;
			final int row = Integer.parseInt(refMatcher.group(2)) - 1;
			setSelector(spreadsheet.getSelector() + String.format(CELL, row, column));
		} else {
			throw new NullPointerException("cell reference format is wrong, it should be \"C11\".");
		}
	}
	
	public boolean isMerged() {
		return getProperty("merid") != null;
	}
	
	private int toColumn(String refference) {
		char[] array = refference.toCharArray();
		double result = 0;
		for (int i = array.length - 1; i >= 0; --i) {
			result = ((array[i] - 64) * Math.pow(26, i) + result); 
		}
		return (int) result;
	}
}
