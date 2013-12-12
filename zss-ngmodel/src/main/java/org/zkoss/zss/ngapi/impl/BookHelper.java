/* BookHelper.java

	Purpose:
		
	Description:
		
	History:
		Mar 24, 2010 5:42:58 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.ngapi.impl;

import java.util.Map;
import java.util.logging.Logger;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;
import org.zkoss.poi.hssf.usermodel.*;
import org.zkoss.poi.hssf.util.*;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.xssf.model.ThemesTable;
import org.zkoss.poi.xssf.usermodel.*;

/**
 * copies from ZSS project
 */
public final class BookHelper {
	public static final String AUTO_COLOR = "AUTO_COLOR";

	//@see #setBorders()
	public static final short BORDER_EDGE_BOTTOM		= 0x01;
	public static final short BORDER_EDGE_RIGHT			= 0x02;
	public static final short BORDER_EDGE_TOP			= 0x04;
	public static final short BORDER_EDGE_LEFT			= 0x08;
	public static final short BORDER_INSIDE_HORIZONTAL	= 0x10;
	public static final short BORDER_INSIDE_VERTICAL 	= 0x20;
	public static final short BORDER_DIAGONAL_DOWN		= 0x40;
	public static final short BORDER_DIAGONAL_UP		= 0x80;
	
	public static final short BORDER_FULL		= 0xFF;
	public static final short BORDER_OUTLINE	= 0x0F;
	public static final short BORDER_INSIDE		= 0x30;
	public static final short BORDER_DIAGONAL	= 0xC0;
	
	//@see #sort()
	public static final int SORT_NORMAL_DEFAULT = 0;
	public static final int SORT_TEXT_AS_NUMBERS = 1;
	public static final int SORT_HEADER_NO  = 0;
	public static final int SORT_HEADER_YES = 1;	
	
	//inner pasteType for #paste & #fill
	public final static int INNERPASTE_NUMBER_FORMATS = 0x01;
	public final static int INNERPASTE_BORDERS = 0x02;
	public final static int INNERPASTE_OTHER_FORMATS = 0x04;
	public final static int INNERPASTE_VALUES = 0x08;
	public final static int INNERPASTE_FORMULAS = 0x10;
	public final static int INNERPASTE_VALUES_AND_FORMULAS = INNERPASTE_FORMULAS + INNERPASTE_VALUES;
	public final static int INNERPASTE_COMMENTS = 0x20;
	public final static int INNERPASTE_VALIDATION = 0x40;
	public final static int INNERPASTE_COLUMN_WIDTHS = 0x80;
	public final static int INNERPASTE_FORMATS = INNERPASTE_NUMBER_FORMATS + INNERPASTE_BORDERS + INNERPASTE_OTHER_FORMATS;
	
	private final static int INNERPASTE_FILL_COPY = INNERPASTE_FORMATS + INNERPASTE_VALUES_AND_FORMULAS + INNERPASTE_VALIDATION;   
	private final static int INNERPASTE_FILL_VALUE = INNERPASTE_VALUES_AND_FORMULAS + INNERPASTE_VALIDATION;   
	private final static int INNERPASTE_FILL_FORMATS = INNERPASTE_FORMATS;   

	//inner pasteOp for #paste & #fill
	public final static int PASTEOP_ADD = 1;
	public final static int PASTEOP_SUB = 2;
	public final static int PASTEOP_MUL = 3;
	public final static int PASTEOP_DIV = 4;
	public final static int PASTEOP_NONE = 0;

	//inner fillType for #fill
	public final static int FILL_DEFAULT = 0x01; //system determine
	public final static int FILL_FORMATS = 0x02; //formats only
	public final static int FILL_VALUES = 0x04; //value+formula+validation+hyperlink (no comment)
	public final static int FILL_COPY = 0x06; //value+formula+validation+hyperlink, formats
	public final static int FILL_DAYS = 0x10;
	public final static int FILL_WEEKDAYS = 0x20;
	public final static int FILL_MONTHS = 0x30;
	public final static int FILL_YEARS = 0x40;
	public final static int FILL_HOURS = 0x50;
	public final static int FILL_GROWTH_TREND = 0x100; //multiplicative relation
	public final static int FILL_LINER_TREND = 0x200; //additive relation
	public final static int FILL_SERIES = FILL_LINER_TREND;
	
	private static final Logger logger = Logger.getLogger(BookHelper.class.getName());
	/*
	 * Returns the associated #rrggbb HTML color per the given POI Color.
	 * @return the associated #rrggbb HTML color per the given POI Color.
	 */
	public static String colorToHTML(Workbook book, Color color) {
		if (book instanceof HSSFWorkbook) {
			return HSSFColorToHTML((HSSFWorkbook) book, (HSSFColor) color);
		} else {
			return XSSFColorToHTML((XSSFWorkbook) book, (XSSFColor) color);
		}
	}
	public static String colorToBorderHTML(Workbook book, Color color) {
		String htmlColor = colorToHTML(book,color);
		if(AUTO_COLOR.equals(htmlColor)){
			return "#000000";
		}
		return htmlColor;
	}
	public static String colorToBackgroundHTML(Workbook book, Color color) {
		String htmlColor = colorToHTML(book,color);
		if(AUTO_COLOR.equals(htmlColor)){
			return "#ffffff";
		}
		return htmlColor;
	}
	public static String colorToForegroundHTML(Workbook book, Color color) {
		String htmlColor = colorToHTML(book,color);
		if(AUTO_COLOR.equals(htmlColor)){
			return "#000000";
		}
		return htmlColor;
	}
	private static byte[] getRgbWithTint(byte[] rgb, double tint) {
		int k = rgb.length > 3 ? 1 : 0; 
		final byte red = rgb[k++];
		final byte green = rgb[k++];
		final byte blue = rgb[k++];
		final double[] hsl = rgbToHsl(red, green, blue);
		final double hue = hsl[0];
		final double sat = hsl[1];
		final double lum = tint(hsl[2], tint);
		return hslToRgb(hue, sat, lum);
	}
	private static double[] rgbToHsl(byte red, byte green, byte blue){
		final double r = (red & 0xff) / 255d;
		final double g = (green & 0xff) / 255d;
		final double b = (blue & 0xff) / 255d;
		final double max = Math.max(Math.max(r, g), b);
		final double min = Math.min(Math.min(r, g), b);
		double h = 0d, s = 0d, l = (max + min) / 2d;
		if (max == min) {
			h = s = 0d; //gray scale
		} else {
			final double d = max - min;
			s = l > 0.5 ? d / (2d - max - min) : d / (max + min);
			if (max == r) {
				h = (g - b) / d + (g < b ? 6d : 0d);
			} else if (max == g) {
				h = (b - r) / d + 2d; 
			} else if (max == b) {
				h = (r - g) / d + 4d;
			}
			h /= 6d;
		}
		return new double[] {h, s, l};
	}
	private static byte[] hslToRgb(double hue, double sat, double lum){
		 double r, g, b;
		 if(sat == 0d){
			 r = g = b = lum; // gray scale
		 } else {
			 final double q = lum < 0.5d ? lum * (1d + sat) : lum + sat - lum * sat;
			 final double p = 2d * lum - q;
			 r = hue2rgb(p, q, hue + 1d/3d);
			 g = hue2rgb(p, q, hue);
			 b = hue2rgb(p, q, hue - 1d/3d);
		 }
		 final byte red = (byte) (r * 255d);
		 final byte green = (byte) (g * 255d);
		 final byte blue = (byte) (b * 255d);
		 return new byte[] {red, green, blue};
	}
	private static double hue2rgb(double p, double q, double t) {
		if(t < 0d) t += 1d;
		if(t > 1d) t -= 1d;
		if(t < 1d/6d) 
			return p + (q - p) * 6d * t;
		if(t < 1d/2d) 
			return q;
		if(t < 2d/3d) 
			return p + (q - p) * (2d/3d - t) * 6d;
		return p;
	}
	private static double tint(double val, double tint) {
		return tint > 0d ? val * (1d - tint) + tint : val * (1d + tint);
	}

	private static String XSSFColorToHTML(XSSFWorkbook book, XSSFColor color) {
		if (color != null) {
			final CTColor ctcolor = color.getCTColor();
			if (ctcolor.isSetIndexed()) {
				byte[] rgb = IndexedRGB.getRGB(color.getIndexed());
				if (rgb != null) {
					return "#"+ toHex(rgb[0])+ toHex(rgb[1])+ toHex(rgb[2]);
				}
			}
			if (ctcolor.isSetRgb()) {
				byte[] argb = ctcolor.isSetTint() ?
					getRgbWithTint(color.getRgb(), color.getTint())/*color.getRgbWithTint()*/ : color.getRgb();
				return argb.length > 3 ? 
					"#"+ toHex(argb[1])+ toHex(argb[2])+ toHex(argb[3])://ignore alpha
					"#"+ toHex(argb[0])+ toHex(argb[1])+ toHex(argb[2]); 
			}
			if (ctcolor.isSetTheme()) {
			    ThemesTable theme = book.getTheme();
			    if (theme != null) {
			    	XSSFColor themecolor = theme.getThemeColor(color.getTheme());
			    	if (themecolor != null) {
			    		if (ctcolor.isSetTint()) {
			    			themecolor.setTint(ctcolor.getTint());
			    		}
			    		return XSSFColorToHTML(book, themecolor); //recursive
			    	}
			    }
			}
		}
	    return AUTO_COLOR;
 	}
	
	private static String HSSFColorToHTML(HSSFWorkbook book, HSSFColor color) {
		return color == null || HSSFColor.AUTOMATIC.getInstance().equals(color) ? AUTO_COLOR : 
			color.isIndex() ? HSSFColorIndexToHTML(book, color.getIndex()) : HSSFColorToHTML((HSSFColorExt)color); 
	}
	private static String HSSFColorToHTML(HSSFColorExt color) {
		short[] triplet = color.getTriplet();
		byte[] argb = new byte[3];
		for (int j = 0; j < 3; ++j) {
			argb[j] = (byte) triplet[j];
		}
		if (color.isTint()) {
			argb = getRgbWithTint(argb, color.getTint());
		}
		return 	"#"+ toHex(argb[0])+ toHex(argb[1])+ toHex(argb[2]); 
	}
	
	private static String HSSFColorIndexToHTML(HSSFWorkbook book, int index) {
		HSSFPalette palette = book.getCustomPalette();
		HSSFColor color = null;
		if (palette != null) {
			color = palette.getColor(index);
		}
		short[] triplet = null;
		if (color != null)
			triplet =  color.getTriplet();
		else {
			final Map<Integer, HSSFColor> colors = HSSFColor.getIndexHash();
			color = colors.get(Integer.valueOf(index));
			if (color != null)
				triplet = color.getTriplet();
		}
		return triplet == null ? null : 
			HSSFColor.AUTOMATIC.getInstance().equals(color) ? AUTO_COLOR : tripletToHTML(triplet);
	}
	public static String toHex(int num) {
		num = num & 0xff;
		final String hex = Integer.toHexString(num);
		return hex.length() == 1 ? "0"+hex : hex; 
	}
	
	private static String tripletToHTML(short[] triplet) {
		return triplet == null ? null : "#"+ toHex(triplet[0])+ toHex(triplet[1])+ toHex(triplet[2]);
	}
}
