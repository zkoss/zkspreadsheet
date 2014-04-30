/*
{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 23, 2014, Created by henrichen
}}IS_NOTE

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.util.Locale;

import org.zkoss.poi.ss.format.Formatters;
import org.zkoss.poi.ss.usermodel.Cell;

/**
 * @author henrichen
 *
 */
public class NumberInputMaskImpl implements org.zkoss.zss.model.NumberInputMask {
	@Override
	public Object[] parseNumberInput(String txt, Locale locale) {
		final char dot = Formatters.getDecimalSeparator(locale);
		final char comma = Formatters.getGroupingSeparator(locale);
		String txt0 = txt;
		if (dot != '.' || comma != ',') {
			final int dotPos = txt.lastIndexOf(dot);
			txt0 = txt.replace(comma, ',');
			if (dotPos >= 0) {
				txt0 = txt0.substring(0, dotPos)+'.'+txt0.substring(dotPos+1);
			}
		}
		
		try {
			final Double val = Double.parseDouble(txt0);
			return new Object[] {new Integer(Cell.CELL_TYPE_NUMERIC), val}; //double
		} catch (NumberFormatException ex) {
			return new Object[] {txt, null};
		}
	}
}
