/**
 * 
 */
package org.zkoss.zss.model.impl;

import static org.junit.Assert.assertEquals;

import java.util.Locale;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.ss.util.NumberToGeneralTextConverter;


/**
 * ZSS-628
 * Test converstion of double numbers to text in "General" format
 * @author henri
 *
 */
public class ToGeneralTextTest {
	//when column width is 12 characters
	private static final Object VAL12[][] = {
		{-38277.429999999935, "-38277.43"},
		{-38277.430000999935, "-38277.43"},
		{-38277.430009999935, "-38277.43001"},
		{0.000123456789012345, "0.000123457"},
		{-0.000123456789012345, "-0.000123457"},
		{0.0000123456789012345, "1.23457E-05"}, // 5
		{-0.0000123456789012345, "-1.23457E-05"},
		{0.00001234, "0.00001234"},
		{0.000012345, "0.000012345"},
		{0.0000123456, "1.23456E-05"},
		{-1.23456789E+112, "-1.2346E+112"}, //10
		{-1.23456789E-112, "-1.2346E-112"},
		{-1.2E-112, "-1.2E-112"},
		{-12345678901., "-12345678901"},
		{-123456789012., "-1.23457E+11"},
		{-1234567890123., "-1.23457E+12"}, //15
		{-12345678901234., "-1.23457E+13"},
		{-123456789012345., "-1.23457E+14"},
		{-1234567890123456., "-1.23457E+15"},
		{-12345678901234567., "-1.23457E+16"},
		{-123456789012345678901., "-1.23457E+20"},
		{-1.23456789012345E+21, "-1.23457E+21"},
	};

	//when column width is 11 characters
	private static final Object VAL11[][] = {
		{-38277.429999999935, "-38277.43"},
		{-38277.430000999935, "-38277.43"},
		{-38277.430009999935, "-38277.43"},
		{0.000123456789012345, "0.000123457"},
		{-0.000123456789012345, "-0.00012346"},
		{0.0000123456789012345, "1.23457E-05"}, //5
		{-0.0000123456789012345, "-1.2346E-05"},
		{0.00001234, "0.00001234"},
		{0.000012345, "0.000012345"},
		{0.0000123456, "1.23456E-05"},
		{-1.23456789E+112, "-1.235E+112"}, //10
		{-1.23456789E-112, "-1.235E-112"},
		{-12345678901., "-1.2346E+10"},
		{-123456789012., "-1.2346E+11"},
		{-1234567890123., "-1.2346E+12"},
		{-12345678901234., "-1.2346E+13"},
		{-123456789012345., "-1.2346E+14"},
		{-1234567890123456., "-1.2346E+15"},
		{-12345678901234567., "-1.2346E+16"},
		{-123456789012345678901., "-1.2346E+20"},
		{-1.23456789012345E+21, "-1.2346E+21"},
	};
	
	//when column width is 6 characters
	private static final Object VAL6[][] = {
		{-38277.429999999935, "-38277"},
		{-38277.430000999935, "-38277"},
		{-38277.430009999935, "-38277"},
		{0.000123456789012345, "0.0001"},
		{-0.000123456789012345, "-1E-04"},
		{0.0000123456789012345, "1E-05"}, //5
		{-0.0000123456789012345, "-1E-05"},
		{0.00001234, "1E-05"},
		{0.000012345, "1E-05"},
		{0.0000123456, "1E-05"},
		{-1.23456789E+112, "######"}, //10
		{-1.23456789E-112, "######"},
		{-12345678901., "-1E+10"},
		{-123456789012., "-1E+11"},
		{-1234567890123., "-1E+12"},
		{-12345678901234., "-1E+13"},
		{-123456789012345., "-1E+14"},
		{-1234567890123456., "-1E+15"},
		{-12345678901234567., "-1E+16"},
		{-123456789012345678901., "-1E+20"},
		{-1.23456789012345E+21, "-1E+21"},
	};
	
	//when column width is 6 characters
	private static final Object VAL3[][] = {
		{-38277.429999999935, "###"},
		{-38277.430000999935, "###"},
		{-38277.430009999935, "###"},
		{0.000123456789012345, "0"},
		{-0.000123456789012345, "-0"},
		{0.0000123456789012345, "0"}, //5
		{-0.0000123456789012345, "-0"},
		{0.00001234, "0"},
		{0.000012345, "0"},
		{0.0000123456, "0"},
		{-1.23456789E+112, "###"}, //10
		{-1.23456789E-112, "-0"},
		{-0.9, "-1"},
		{0.9, "0.9"},
		{0.0923456789012345, "0.1"},
		{-0.0923456789012345, "-0"},
	};

	/**
	 * @throws java.lang.Exception
	 */
	@Before
	public void setUp() throws Exception {
	}

	/**
	 * @throws java.lang.Exception
	 */
	@After
	public void tearDown() throws Exception {
	}

	@Test
	public void testParseNumberInputWidth12() {
		for(int j= 0; j < VAL12.length; ++j) {
			String result = NumberToGeneralTextConverter.toGeneralText(((Number)VAL12[j][0]).doubleValue(), Locale.US, 12);
			assertEquals(VAL12[j][0]+"->"+j, VAL12[j][1], result);
		}
	}
	
	@Test
	public void testParseNumberInputWidth11() {
		for(int j= 0; j < VAL11.length; ++j) {
			String result = NumberToGeneralTextConverter.toGeneralText(((Number)VAL11[j][0]).doubleValue(), Locale.US, 11);
			assertEquals(VAL11[j][0]+"->"+j, VAL11[j][1], result);
		}
	}

	@Test
	public void testParseNumberInputWidth6() {
		for(int j= 0; j < VAL6.length; ++j) {
			try {
			String result = NumberToGeneralTextConverter.toGeneralText(((Number)VAL6[j][0]).doubleValue(), Locale.US, 6);
			assertEquals(VAL6[j][0]+"->"+j, VAL6[j][1], result);
			} catch(Exception e) {
				System.out.println("Exception:"+ VAL6[j][0]+"->"+j +"\n"+e);
				break;
			}
		}
	}

	@Test
	public void testParseNumberInputWidth3() {
		for(int j= 0; j < VAL3.length; ++j) {
			try {
			String result = NumberToGeneralTextConverter.toGeneralText(((Number)VAL3[j][0]).doubleValue(), Locale.US, 3);
			assertEquals(VAL3[j][0]+"->"+j, VAL3[j][1], result);
			} catch(Exception e) {
				System.out.println("Exception:"+ VAL3[j][0]+"->"+j +"\n"+e);
				break;
			}
		}
	}
}
