package org.zkoss.zss;

import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import org.apache.commons.math.complex.Complex;
import org.junit.Assert;
import org.zkoss.poi.ss.formula.functions.ComplexFormat;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SSheet;

public class AssertUtil {
	
	public static double DELTA = 1E-8; // formula precision setting : 10 ^ -8
	
	/**
	 * test whether two complex is equals 
	 */
	public static void assertComplexEquals(String complex1, String complex2) {
		
		try {
			
			complex1 = replaceiWith1i(complex1);
			complex2 = replaceiWith1i(complex2);
			
			ComplexFormat cf = new ComplexFormat();
			Complex c1 = cf.parse(complex1);
			Complex c2 = cf.parse(complex2);
			
			double real1 = c1.getReal();
			double im1 = c1.getImaginary();
			double real2 = c2.getReal();
			double im2 = c2.getImaginary();

			if(Math.abs(real1 - real2) <= DELTA && Math.abs(im1 - im2) <= DELTA) {
				return;
			}
			
		} catch (ParseException e) {
			e.printStackTrace();
		}
		
		org.junit.Assert.fail(complex1 + " is not equals to " + complex2);;

	}
	
	/**
	 * only accept i, j, k as imaginary character
	 */
	private static String getComplexImaginaryChar(String complex) {
		String s = "i";
		int i = complex.indexOf(s);
		if (i == -1) {
			s = "j";
			i = complex.indexOf(s);
			if (i == -1) {
				s = "k";
				i = complex.indexOf(s);
				if (i == -1) {
					throw new RuntimeException("only accept i, j, k as imaginary character");
				}
				return s; // k
			}
			return s; // j
		}
		return s; // i
	}
	
	/**
	 * Apache common cannot parse complex like "8+i", must replace i as 1i.
	 * So the result string will become "8+1i".
	 * Only replace when the (indexOf "i" - 1) is empty or operator.
	 * There are two example case:
	 * 8+i, i.  
	 */
	private static String replaceiWith1i(String complex) {
		
		String imChar = getComplexImaginaryChar(complex);
		
		if(complex.indexOf(imChar) == 0) {
			return complex.replace(imChar, "1" + imChar);
		}
		
		if(complex.charAt(complex.indexOf(imChar) - 1) == '+' || complex.charAt(complex.indexOf(imChar) - 1) == '-') {
			return complex.replace(imChar, "1" + imChar);
		}
		
		return complex;
	}
	
	
	public static void assertMeregedRegion(String[] regions,Sheet sheet){
		Assert.assertEquals(regions.length, sheet.getInternalSheet().getNumOfMergedRegion());
		List<String> expect = new ArrayList<String>(Arrays.asList(regions));
		List<String> result = new ArrayList<String>();
		for(int i=0;i<regions.length;i++){
			String fm = sheet.getInternalSheet().getMergedRegion(i).getReferenceString();
			result.add(fm);
		}
		Collections.sort(expect);
		Collections.sort(result);
		Assert.assertEquals(expect.toString(), result.toString());
	}
	
	public static void assertMergedRange(Range range) {
		if(!isMergedRange(range)){
			Assert.fail(range.asString() + " is not a merged range");
		}
	}
	
	public static void assertNotMergedRange(Range range) {
		if(isMergedRange(range)){
			Assert.fail(range.asString() + " is a merged range");
		}
	}
	
	private static boolean isMergedRange(Range range){
		SSheet sheet = range.getSheet().getInternalSheet();
		// go through all region
		for (int number = sheet.getNumOfMergedRegion(); number > 0; number--) {
			CellRegion addr = sheet.getMergedRegion(number - 1);
			// match four corner
			if (addr.getRow() == range.getRow() && addr.getLastRow() == range.getLastRow() && addr.getColumn() == range.getColumn() && addr.getLastColumn() == range.getLastColumn()) {
				return true;
			}
		}
		return false;
	}

	public static void assertLeftBorder(Range range,BorderType type,String color) {
		CellStyle style = range.getCellStyle();
		String msg = "at "+range.toCellRange(0, 0).asString()+" left";
		BorderType tbt = style.getBorderLeft();
		String tcolor = style.getBorderLeftColor().getHtmlColor();
		if(!type.equals(BorderType.NONE)){
			if(tbt.equals(BorderType.NONE)){
				style = range.toCellRange(0, -1).getCellStyle(); 
				tbt = style.getBorderRight();
				tcolor = style.getBorderRightColor().getHtmlColor();
			}
		}
		Assert.assertEquals(msg, type, tbt);
		if(color!=null){
			Assert.assertEquals(msg, color, tcolor);
		}
	}
	public static void assertRightBorder(Range range,BorderType type,String color) {
		CellStyle style = range.getCellStyle();
		String msg = "at "+range.toCellRange(0, 0).asString()+" right";
		BorderType tbt = style.getBorderRight();
		String tcolor = style.getBorderRightColor().getHtmlColor();
		if(!type.equals(BorderType.NONE)){
			if(tbt.equals(BorderType.NONE)){
				style = range.toCellRange(0, 1).getCellStyle(); 
				tbt = style.getBorderLeft();
				tcolor = style.getBorderLeftColor().getHtmlColor();
			}
		}
		Assert.assertEquals(msg, type, tbt);
		if(color!=null){
			Assert.assertEquals(msg, color, tcolor);
		}
		
	}
	public static void assertTopBorder(Range range,BorderType type,String color) {
		CellStyle style = range.getCellStyle();
		String msg = "at "+range.toCellRange(0, 0).asString()+" top";
		BorderType tbt = style.getBorderTop();
		String tcolor = style.getBorderTopColor().getHtmlColor();
		if(!type.equals(BorderType.NONE)){
			if(tbt.equals(BorderType.NONE)){
				style = range.toCellRange(-1, 0).getCellStyle(); 
				tbt = style.getBorderBottom();
				tcolor = style.getBorderBottomColor().getHtmlColor();
			}
		}
		Assert.assertEquals(msg, type, tbt);
		if(color!=null){
			Assert.assertEquals(msg, color, tcolor);
		}
	}
	public static void assertBottomBorder(Range range,BorderType type,String color) {
		CellStyle style = range.getCellStyle();
		String msg = "at "+range.toCellRange(0, 0).asString()+" bottom";
		BorderType tbt = style.getBorderBottom();
		String tcolor = style.getBorderBottomColor().getHtmlColor();
		if(!type.equals(BorderType.NONE)){
			if(tbt.equals(BorderType.NONE)){
				style = range.toCellRange(1, 0).getCellStyle(); 
				tbt = style.getBorderTop();
				tcolor = style.getBorderTopColor().getHtmlColor();
			}
		}
		Assert.assertEquals(msg, type, tbt);
		if(color!=null){
			Assert.assertEquals(msg, color.toLowerCase(), tcolor.toLowerCase());
		}
	}

}
