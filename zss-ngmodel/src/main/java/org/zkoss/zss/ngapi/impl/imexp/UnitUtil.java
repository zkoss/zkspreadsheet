/* UnitUtil.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;


/**
 * copied from ZSS. We should remove it after integration.
 * Pure unit conversion methods.
 */
public class UnitUtil {

	/** convert pixel to point */
	public static int pxToPoint(int px) {
		return px * 72 / 96; //assume 96dpi
	}
	
	/** convert point to pixel */
	public static int pointToPx(int point) {
		return point * 96 / 72; //assume 96dpi
	}
	
	
	/** convert pixel to EMU */
	public static int pxToEmu(int px) {
		//refer form ActionHandler,ChartHelper
		return (int) Math.round(((double)px) * 72 * 20 * 635 / 96); //assume 96dpi
	}
	
	/** convert EMU to pixel, 1 twip == 635 emu */
	public static int emuToPx(int emu) {
		//refer form ChartHelper
		return (int) Math.round(((double)emu) * 96 / 72 / 20 / 635); //assume 96dpi
	}

	/** convert twip (1/20 point) to pixel */
	public static int twipToPx(int twip) {
		return twip * 96 / 72 / 20; //assume 96dpi
	}
	
	/** convert pixel to twip (1/20 point) */
	public static int pxToTwip(int px) {
		return px * 72 * 20 / 96; //assume 96dpi
	}

	/** convert file 1/256 character width to pixel */
	public static int fileChar256ToPx(int char256, int charWidth) {
		final double w = (double) char256;
		return (int) Math.floor(w * charWidth / 256 + 0.5);
	}

	/** convert pixel to file 1/256 character width */
	public static int pxToFileChar256(int px, int charWidth) {
		final double w = (double) px;
		return (int) Math.floor(w * 256 / charWidth + 0.5);
	}

	/** 
	 * Convert default columns width (in character) to pixel.
	 * 5 pixels are margin padding.(There are 4 pixels of margin padding (two on each side), plus 1 pixel padding for the gridlines.)
	 * Description of how column widths are determined in Excel (http://support.microsoft.com/kb/214123)
	 * @param columnWidth number of character
	 * @param charWidth Using the Calibri font, the maximum digit width of 11 point font size is 7 pixels (at 96 dpi).
	 * @return width in pixel  Rounds this number up to the nearest multiple of 8 pixels.
	 */ 
	public static int defaultColumnWidthToPx(int columnWidth, int charWidth) {
		final int w = columnWidth * charWidth + 5;
		final int diff = w % 8;
		return w + (diff > 0 ? (8 - diff) : 0);
	}

	/**
	 * Convert default column width in pixel to number of character by reverse defaultColumnWidthToPx() and ignore the mod(%) operation.
	 * @param px default column width in pixel
	 * @param charWidth  one character width in pixel of normal style's font 
	 * @return default column width in character
	 */
	public static int pxToDefaultColumnWidth(int px, int charWidth) {
		return (px - 5) / charWidth;
	}

	// Formula: Pixels = Inches * DPI, assume 96 DPI
	public static int incheToPx(double inches) {
		return (int) (inches * 96);
	}

	// Formula: Inches = Pixels / DPI, assume 96 DPI
	public static double pxToInche(int px) {
		return px / 96.0;
	}

	public static double pxToCTChar(int px, int charWidth) {
		return (double) pxToFileChar256(px, charWidth) / 256;
	}

}
