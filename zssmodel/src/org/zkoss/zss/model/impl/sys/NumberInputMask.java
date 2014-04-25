/*
{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 23, 2014, Created by henri
}}IS_NOTE

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/

package org.zkoss.zss.model.impl.sys;

import java.util.Locale;

import org.zkoss.poi.ss.format.Formatters;
import org.zkoss.zss.model.SCell.CellType;

/**
 * Responsible for number input mask. 
 * 
 * @author henri
 * @since 3.5.0
 */
public class NumberInputMask {
	/**
	 * @param txt the input text
	 */
	public Object[] parseNumberInput(String txt, Locale locale) {
		final Object[] aStringResult = new Object[] {txt, null};
		final String txt0 = txt.trim();
		int len = txt0.length();
		if (txt == null || len == 0) {
			return aStringResult;
		}
    
		final char dot = Formatters.getDecimalSeparator(locale);
		final char comma = Formatters.getGroupingSeparator(locale);
		final String currency = Formatters.getCurrencySymbol(locale);
		
		
		StringBuilder sb = new StringBuilder(len);
		int commaPos = -17; //distance must >= 3
		boolean percent = false;
		int parenPos = -1;
		boolean digit = false;
		boolean decimal = false; 
		boolean withCurrency = false;
		boolean exp = false;
		boolean negative = false;
		int signPos = -1;
		
		int j = 0;
		//first character
		char c = txt0.charAt(j); 

		if (c == '(') {
			int k = txt0.lastIndexOf(')',  len-1);
			if (k == (len - 1)) { // (xxx) paren
				parenPos = 0;
				--len;
				negative = true;
			} else if (k > 0 && "%".equals(txt0.substring(k+1, len).trim())){ //(123) %
				parenPos = 0;
				len = k;
				negative = true;
				percent = true;
			} else {
				return aStringResult;
			}
		} else if (c == '+' || c == '-') { // + sign
			signPos = 0;
			if (c == '-') {
				negative = !negative;
			}
		} else if (txt0.startsWith(currency)) { //currency
			withCurrency = true;
			j += currency.length() - 1;
		} else if (c == '$') { //$ currency
			withCurrency = true;
		} else if (Character.isDigit(c)) { //digit
			sb.append(c);
			digit = true;
		} else if (c == dot) { //decimal
			sb.append('.');
			decimal = true;
		} else if (c == '%') { //percent
			percent = true;
		} else {
			// none of paren negative, '.', '+', '-', currency, digits => illegal
			return aStringResult;
		}
		
		//2nd character
		c = txt0.charAt(++j); 
		if (parenPos >= 0) { //paren => allow currency, digit, percent
			if (Character.isDigit(c)) { //digit
				sb.append(c);
				digit = true;
			} else if (txt0.startsWith(currency, j)) { //currency
				withCurrency = true;
				j += currency.length() - 1;
			} else if (c == '$') { //$ currency
				withCurrency = true;
			} else if (c == '%') { //%
				percent = true;
			} else {
				return aStringResult;
			}
		} else if (digit) { //digit => allow digit, decimal, comma (must > 3), 'e' or 'E'
			if (c == comma) { //comma
				commaPos = j;
			} else if (c == dot) { //decimal
				sb.append('.');
				decimal = true;
			} else if (Character.isDigit(c)) { //digit
				sb.append(c);
				//digit = true;
			} else if ('e' == c || 'E' == c) { //'e' or 'E'
				sb.append(c);
				exp = true;
			} else {
				return aStringResult;
			}
		} else if (decimal) { //decimal => allow digit only
			if (Character.isDigit(c)) {
				sb.append(c);
				digit = true;
			} else {
				return aStringResult;
			}
		} else if (signPos >= 0) { //sign => allow paren, sign, digit, currency
			if (Character.isDigit(c)) { //digit
				sb.append(c);
				digit = true;
			} else if (c == '(') { //(xxx) paren; NOT a negative
				int k = txt0.lastIndexOf(')',  len-1);
				if (k == (len - 1)) { // (xxx) paren
					parenPos = j;
					--len;
				} else if (k > 0 && "%".equals(txt0.substring(k+1, len).trim())){ //(123) %
					parenPos = j;
					len = k;
					percent = true;
				} else {
					return aStringResult;
				}
			} else if (c == '+' || c == '-') { //sign
				signPos = j;
				if (c == '-') {
					negative = !negative;
				}
			} else if (txt0.startsWith(currency, j)) { //currency
				withCurrency = true;
				j += currency.length() - 1;
			} else if (c == '$') { //$ currency
				withCurrency = true;
			} else {
				return aStringResult;
			}
		} else if (withCurrency) { //currency => allow paren, sign, digit, decimal
			if (Character.isDigit(c)) { //digit
				sb.append(c);
				digit = true;
			} else if (c == dot) { //decimal
				sb.append('.');
				decimal = true;
			} else if (c == '(' && txt0.charAt(len-1) == ')') { //(xxx) paren; a negative
				parenPos = j;
				--len;
				negative = true;
			} else if (c == '+' || c =='-') { //sign
				signPos = j;
				if (c == '-') {
					negative = true;
				}
			} else {
				return aStringResult;
			}
		} else if (percent) { //percent => allow digit, decimal, paren, sign
			if (Character.isDigit(c)) {
				sb.append(c);
			} else if (c == dot) {
				sb.append('.');
				decimal = true;
			} else if (c == '(' && txt0.charAt(len-1) == ')') { //(xxx) paren; a negative
				parenPos = j;
				--len;
				negative = true;
			} else if (c == '-' || c == '+') {
				signPos = j;
				if (c == '-') {
					negative = true;
				}
			} else {
				return aStringResult;
			}
		}
		
		//3rd character and after
		for (++j; j < len; ++j) {
			c = txt0.charAt(j);

			if (Character.isDigit(c)) { //digit
				sb.append(c);
				digit = true;
			} else if (c == '(') { //paren
				if (digit || (parenPos != (j-1) && signPos != (j-1)) || signPos < 0) {
					return aStringResult;
				} else {
					int k = txt0.lastIndexOf(')',  len-1);
					if (k == (len - 1)) { // (xxx) paren
						parenPos = j;
						--len;
					} else if (k > 0 && !percent && "%".equals(txt0.substring(k+1, len).trim())){ //+((123)%)
						parenPos = j;
						len = k;
						percent = true;
					} else {
						return aStringResult;
					}
				}
			} else if (c == dot) { //decimal
				if (!decimal && !exp && (j - commaPos) > 3) {
					decimal = true;
					sb.append('.');
				} else {
					return aStringResult;
				}
			} else if (c == 'E' || c == 'e') { //exp
				if (digit && !exp && (j - commaPos) > 3) {
					exp = true;
					sb.append(c);
				} else {
					return aStringResult;
				}
			} else if (c == '%') { //percent
				if (!percent && digit && !withCurrency) {
					percent = true;
				} else {
					return aStringResult;
				}
			} else if (c == comma) { //comma
				if (digit && (j - commaPos) > 3 && !exp && (parenPos < 0 || withCurrency || signPos < 0)) {
					commaPos = j;
				} else {
					return aStringResult;
				}
			} else if (c == '-' || c == '+') { //sign
				if (parenPos < 0 && signPos == (j-1)) {
					signPos = j;
					if (c == '-') {
						negative = !negative;
					}
				} else {
					return aStringResult;
				}
			} else {
				return aStringResult;
			}
		}
		
		// distance between thousand separators must > 3
		if ((len - commaPos) < 4) { 
			return aStringResult;
		}

		try {
			Double val = Double.parseDouble(sb.toString());
			//prepare format
			String format = null;
			if (exp) { // Strongest format
				format = "0.00E+00";
			} else if (percent) {
				format = decimal ? "0.00%" : "0%";
			} else if (withCurrency) {
				format = decimal ? "$#,##0.00_);[Red]($#,##0.00)" : "$#,##0_);[Red]($#,##0)"; 
			} else if (commaPos > 0) {
				format = decimal ? "#,##0.00" : "#,##0"; 
			}
			if (percent) {
				val /= 100;
			}
			if (negative) {
				val = -val;
			}
			return format == null ? 
					new Object[] {CellType.NUMBER, val} :
					new Object[] {CellType.NUMBER, val, format}; //double
		} catch (NumberFormatException ex) {
			return aStringResult;
		}
	}
}
