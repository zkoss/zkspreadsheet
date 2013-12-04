package org.zkoss.zss.ngmodel.impl.sys;

import java.util.Locale;

import org.zkoss.poi.ss.format.Formatters;
import org.zkoss.poi.ss.usermodel.ErrorConstants;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;

public class InputConverter {
	static private DateInputMask dateInputMask = new DateInputMask();
	/*
	 * TODO could cell be null?
	 */
	public static InputResult editTextToValue(String text, NCell cell, InputParseContext context) {
		InputResultImpl result = null;
		if (text != null) {
			final String formatStr = cell == null ? 
					null : cell.getCellStyle().getDataFormat();
			Object[] convertedResult = editTextToValue(text, formatStr, context.getLocale());
			result = new InputResultImpl();
			result.setType((CellType)convertedResult[0]);
			result.setValue(convertedResult[1]);
		}
		return result;
		
	}
	
	/**
	 * 
	 * @param txt the text to be input
	 * @param formatStr the cell text format 
	 * @return object array with the value type in 0(an Integer), 
	 * 		the value in 1(an Object), and the date format in 2(a String if parse as a date)   
	 */
	public static Object[] editTextToValue(String txt, String formatStr, Locale locale) {
		if (txt != null) {
			//bug #300:	Numbers in Text-cells are not treated as text (leading zero is removed)
			if (formatStr != null) {
				if (isStringFormat(formatStr)) { 
					return new Object[] {CellType.STRING, txt}; //string
				}
			}
			if (txt.startsWith("=")) {
				if (txt.trim().length() > 1) {
					return new Object[] {CellType.FORMULA, txt.substring(1)}; //formula 
				} else {
					return new Object[] {CellType.STRING, txt}; //string
				}
			} else if ("true".equalsIgnoreCase(txt) || "false".equalsIgnoreCase(txt)) {
				return new Object[] {CellType.BOOLEAN, Boolean.valueOf(txt)}; //boolean
			} else if (txt.startsWith("#")) { //might be an error
				final byte err = getErrorCode(txt);
				return err < 0 ? 
					new Object[] {CellType.STRING, txt}: //string
					new Object[] {CellType.ERROR, new Byte(err)}; //error
			} else {
	            return parseEditTextToDoubleDateOrString(txt, locale); //ZSS-67
			}
		}
		return null;
	}
	
	private static Object[] parseEditTextToDoubleDateOrString(String txt, Locale locale) {
		//TODO prepare a NumberInputMask that will set number format if input with comma thousand separator.
//		final Locale locale = ZssContext.getCurrent().getLocale(); //ZSS-67
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
			return new Object[] {CellType.NUMBER, val}; //double
		} catch (NumberFormatException ex) {
			return parseEditTextToDateOrString(txt, locale);
		}
	}
	
	private static Object[] parseEditTextToDateOrString(String txt, Locale locale) {
		final Object[] results = dateInputMask.parseDateInput(txt, locale); 
		if (results[0] instanceof String) { 
			return new Object[] {CellType.STRING, results[0]}; //string
		} else { //if (result[0] instanceof Date)
			return new Object[] {CellType.NUMBER, results[0], results[1]}; //date with format
		}
	}

	private static byte getErrorCode(String errString) {
		if ("#NULL!".equals(errString))
			return ErrorConstants.ERROR_NULL;
		if ("#DIV/0!".equals(errString))
			return ErrorConstants.ERROR_DIV_0;
		if ("#VALUE!".equals(errString))
			return ErrorConstants.ERROR_VALUE;
		if ("#REF!".equals(errString))
			return ErrorConstants.ERROR_REF;
		if ("#NAME?".equals(errString))
			return ErrorConstants.ERROR_NAME;
		if ("#NUM!".equals(errString))
			return ErrorConstants.ERROR_NUM; 
		if ("#N/A".equals(errString))
			return ErrorConstants.ERROR_NA;
		return -1;
	}

	private static boolean isStringFormat(String formatStr) {
		return "@".equals(formatStr); //TODO, shall prepare a regular expression match check!
	}
}
