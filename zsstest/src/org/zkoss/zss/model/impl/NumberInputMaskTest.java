/**
 * 
 */
package org.zkoss.zss.model.impl;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

/**
 * @author henri
 *
 */
public class NumberInputMaskTest {
	NumberInputMask _formater;

	private static final String OK[][] = {
		
		//thousand separator
		{"1,234,5678", "#,##0"},
		{"123,45678", "#,##0"},
		{"(123,4567)", "#,##0"},
		
		{"1,234,567.0", "#,##0.00"},
		{"1,234,567.12", "#,##0.00"}, 
		{"12,4567.123", "#,##0.00"},
		
		
		//percent without decimal
		{"%123", "0%"},
		{"123%", "0%"},
		
		{"%(123)", "0%"},
		{"(123)%", "0%"},
		{"(%123)", "0%"},
		{"(123%)", "0%"},

		{"+((123%))", "0%"},
		{"+(123)%", "0%"},
		{"+((123)%)", "0%"},

		//percent with decimal
		{"%123.", "0.00%"},
		{"%123.1", "0.00%"},
		{"123.12%", "0.00%"},
		{"123.123%", "0.00%"},

		//parenthesis => negative
		{"(123)", null},
		
		//sign and one parenthesis
		{"-(123)", null},
		{"-+(123)", null},
		{"+-(123)", null},
		{"---(123)", null},
		
		//sign and one parenthesis
		{"+(123)", null},
		{"++(123)", null},
		{"--(123)", null},
		{"-+-(123)", null},
		
		//sign and multiple parenthesis
		{"+((123))", null},
		{"++(((123)))", null},
		{"--((((123))))", null},
		{"-+-(((((123)))))", null},
		
		//currency without decimal
		{"$1234567", "$#,##0_);[Red]($#,##0)"},
		{"$1,234,567", "$#,##0_);[Red]($#,##0)"},
		{"$123,4567", "$#,##0_);[Red]($#,##0)"},
		
		//currency with decimal
		{"$1234567.89", "$#,##0.00_);[Red]($#,##0.00)"},
		{"$1,234,567.89", "$#,##0.00_);[Red]($#,##0.00)"},
		{"$123,4567.89", "$#,##0.00_);[Red]($#,##0.00)"},
		
		//currency with decimal and parenthesis/negative
		{"$(1234567.89)", "$#,##0.00_);[Red]($#,##0.00)"},
		{"-$1234567.89", "$#,##0.00_);[Red]($#,##0.00)"},
		{"($1,234,567.89)", "$#,##0.00_);[Red]($#,##0.00)"}, 
		{"$(123,4567.89)", "$#,##0.00_);[Red]($#,##0.00)"},
		{"-$123,4567.89", "$#,##0.00_);[Red]($#,##0.00)"},
		{"$-123,4567.89", "$#,##0.00_);[Red]($#,##0.00)"},
		
		//scientific
		{"1234567E1", "0.00E+00"},
		{"1,234567E1", "0.00E+00"},
		{"123,4567E1", "0.00E+00"},
		{"1234567E1", "0.00E+00"},
		{"1,234,567E1", "0.00E+00"},
		
		//scientific with currency/parenthesis/percent
		{"(1,234567E1)", "0.00E+00"},
		{"-123,4567E1", "0.00E+00"},
		{"-$1234567E1", "0.00E+00"},
		{"($1,234,567E1)", "0.00E+00"},
		{"$(1,234,567E1)", "0.00E+00"},
		{"(1,234567E1)%", "0.00E+00"},
		{"%(1,234567E1)", "0.00E+00"},
	};
	
	private static final String FAIL[] = {
		//separator distance must >= 3
		"1,234,57", 
		"12,456,81.9", 
		"123,45E10",
		
		//percent with more than one parenthesis 
		"%((123))", 
		"((123))%", 
		"((%123))", 
		"((123%))",
		"((123)%)",
		
		//sign can have multiple parenthesis, but cannot mix with %
		"-(%(123))",
		
		//sign, parenthesis, and thousand separator is not legal
		"+(1,234)", 
		"++(1,234)", 
		"--(1,234)", 
		"-+-(1,234)",
		
		//currency and percent not allowed 
		"$(123)%", 
		"$(123E1)%", 
	};
	
	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
		_formater = new NumberInputMask();
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
		_formater = null;
	}

	@Test
	public void testParseNumberInput() {
		for(int j= 0; j < OK.length; ++j) {
			Object[] result = _formater.parseNumberInput(OK[j][0], Locale.US);
			assertTrue(OK[j][0]+"->"+j,  result[1] != null);
			assertEquals(OK[j][0]+"->"+j, OK[j][1], result.length > 2 ? (String) result[2] : null);
		}
		for(int j= 0; j < FAIL.length; ++j) {
			Object[] result = _formater.parseNumberInput(FAIL[j], Locale.US);
			assertFalse(FAIL[j]+"->"+j, result[1] != null);
			assertEquals(FAIL[j]+"->"+j, FAIL[j], (String)result[0]);
		}
	}
}
