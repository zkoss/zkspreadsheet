/**
 * 
 */
package org.zkoss.zss.engine.impl;

import static org.junit.Assert.assertEquals;

import junit.framework.Assert;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.zkoss.poi.ss.util.ToExcelNumberConverter;


/**
 * ZSS-628
 * Test converstion of double numbers to text in "General" format
 * @author henri
 *
 */
public class ToExcelNumberTest {
	//when column width is 12 characters
	private static final Object VAL[][] = {
		{-732870.32, "-732870.32"},
		{-207279.38, "-207279.38"},
		{-4737.63, "-4737.63"},
		{-4737.89, "-4737.89"},
		{-516115.42, "-516115.42"},
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
	public void testParseNumber() {
		double c1 = ((Double) VAL[0][0]).doubleValue();
		double c2 = ((Double) VAL[1][0]).doubleValue();
		double c3 = ((Double) VAL[2][0]).doubleValue();
		double c4 = ((Double) VAL[3][0]).doubleValue();
		double c5 = ((Double) VAL[4][0]).doubleValue();
		
		for(int j= 0; j < VAL.length; ++j) {
			double val = ((Double)VAL[j][0]).doubleValue();
			double valprime = toExcelNumber(val, false);
			assertEquals(""+val+"->"+j, val, valprime, 0.000000000000000001);
			assertEquals(""+val+"->"+j, VAL[j][1], Double.toString(valprime));
		}
		
		double c6 = 0.12345678901234567;
		double c7 = -895326.34 - -857048.91;
		double c8 = -520853.30999999994;
		
		assertEquals(""+c6, 0.123456789012345, toExcelNumber(c6, false), 0.000000000000000001);
		assertEquals(""+c6, 0.123456789012346, toExcelNumber(c6, true), 0.000000000000000001);
		assertEquals(""+c7, -38277.4299999999, toExcelNumber(c7, false), 0.000000000000000001);
		assertEquals(""+c7, -38277.4299999999, toExcelNumber(c7, true), 0.000000000000000001);
		assertEquals(""+c8, -520853.309999999, toExcelNumber(c8, false), 0.000000000000000001);
		assertEquals(""+c8, -520853.31, toExcelNumber(c8, true), 0.000000000000000001);

		double result =  c1 - c2;
		result = toExcelNumber(result, true);
		assertEquals(""+c1+" - "+c2, -525590.94, result, 0.000000000000000001);
		
		result -= c3;
		result = toExcelNumber(result, true);
		assertEquals(""+c1+" - "+c2+" - "+c3, -520853.31, result, 0.000000000000000001);

		result -= c4;
		result = toExcelNumber(result, true);
		assertEquals(""+c1+" - "+c2+" - "+c3+" - "+c4, -516115.42, result, 0.000000000000000001);
		
		result -= c5;
		result = toExcelNumber(result, true);
		assertEquals(""+c1+" - "+c2+" - "+c3+" - "+c4+" - "+c5, 0.0, result, 0.000000000000000001);
	}

	public String toRawBits(double value) {
		long rawBits = Double.doubleToLongBits(value);
		return Long.toHexString(rawBits);
	}
	public double toExcelNumber(double value, boolean rounding) {
		return ToExcelNumberConverter.toExcelNumber(value, rounding);
	}
}