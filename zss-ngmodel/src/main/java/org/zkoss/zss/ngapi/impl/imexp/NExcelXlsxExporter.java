/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import java.io.*;
import java.util.*;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.FillPattern;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NFont.Boldweight;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;
import org.zkoss.zss.ngmodel.NFont.Underline;
/**
 * 
 * @author dennis, kuro
 * @since 3.5.0
 */
public class NExcelXlsxExporter extends AbstractExcelExporter {
	
//	private XSSFWorkbook workbook;
//	private Map<NCellStyle, XSSFCellStyle> styleTable = new HashMap<NCellStyle, XSSFCellStyle>();
//	private Map<NFont, XSSFFont> fontTable = new HashMap<NFont, XSSFFont>();
//	private Map<NColor, XSSFColor> colorTable = new HashMap<NColor, XSSFColor>();
	
	@Override
	public void export(NBook book, OutputStream fos) throws IOException {
		ReadWriteLock lock = book.getBookSeries().getLock();
		lock.readLock().lock();
		
		workbook = new XSSFWorkbook();
		
		// TODO, API isn't available 
		//workbook.setActiveSheet(index);
		//workbook.setFirstVisibleTab(index);
		//workbook.setForceFormulaRecalculation(value);
		//workbook.setHidden(hiddenFlag);
		//workbook.setMissingCellPolicy(missingCellPolicy);
		//workbook.setPrintArea(sheetIndex, reference);
		//workbook.setSelectedTab(index);
		//workbook.setSheetHidden(sheetIx, hidden);
		//workbook.setSheetName(sheetIndex, sheetname);
		//workbook.setSheetOrder(sheetname, pos);
		
		for(NSheet sheet : book.getSheets()) {
			exportSheet(sheet);
		}
		
		exportNamedRange(book);
		
		try{
			workbook.write(fos);
			
		} finally {
			lock.readLock().unlock();
		}
	}
	
	/**
	 * Utility to adapt ZSS 3.5 cell style into XSSFCellStyle
	 * @param cellStyle
	 * @return a XSSFCellStyle Object
	 */
//	protected XSSFCellStyle toXSSFCellStyle(NCellStyle cellStyle) {
//		
//		// instead of creating a new style, use old one if exist
//		XSSFCellStyle xssfCellStyle = (XSSFCellStyle) styleTable.get(cellStyle);
//		if(xssfCellStyle != null) {
//			return xssfCellStyle;
//		}
//		
//		xssfCellStyle = (XSSFCellStyle) workbook.createCellStyle();
//		
//		xssfCellStyle.setAlignment(toPOIAlignment(cellStyle.getAlignment()));
//		
//		/* Bottom Border */
//		xssfCellStyle.setBorderBottom(toPOIBorderType(cellStyle.getBorderBottom()));
//		xssfCellStyle.setBottomBorderColor(toXSSFColor(cellStyle.getBorderBottomColor()));
//		
//		/* Left Border */
//		xssfCellStyle.setBorderLeft(toPOIBorderType(cellStyle.getBorderLeft()));
//		xssfCellStyle.setLeftBorderColor(toXSSFColor(cellStyle.getBorderLeftColor()));
//		
//		/* Right Border */
//		xssfCellStyle.setBorderRight(toPOIBorderType(cellStyle.getBorderRight()));
//		xssfCellStyle.setRightBorderColor(toXSSFColor(cellStyle.getBorderRightColor()));
//		
//		/* Top Border*/
//		xssfCellStyle.setBorderTop(toPOIBorderType(cellStyle.getBorderTop()));
//		xssfCellStyle.setTopBorderColor(toXSSFColor(cellStyle.getBorderTopColor()));
//		
//		/* Fill Foreground Color */
//		xssfCellStyle.setFillForegroundColor(toXSSFColor(cellStyle.getFillColor()));
//		
//		xssfCellStyle.setFillPattern(toPOIFillPattern(cellStyle.getFillPattern()));
//		xssfCellStyle.setVerticalAlignment(toPOIVerticalAlignment(cellStyle.getVerticalAlignment()));
//		xssfCellStyle.setWrapText(cellStyle.isWrapText());
//		xssfCellStyle.setLocked(cellStyle.isLocked());
//		xssfCellStyle.setHidden(cellStyle.isHidden());
//		
//		// refer from BookHelper#setDataFormat
//		XSSFDataFormat df = (XSSFDataFormat) workbook.createDataFormat();
//		short fmt = df.getFormat(cellStyle.getDataFormat());
//		xssfCellStyle.setDataFormat(fmt);
//		
//		// font
//		xssfCellStyle.setFont(toXSSFFont(cellStyle.getFont()));
//		
//		// put into table
//		styleTable.put(cellStyle, xssfCellStyle);
//		
//		return xssfCellStyle;
//	}
//	
//	protected XSSFColor toXSSFColor(NColor color) {
//		XSSFColor xssfFillColor = (XSSFColor) colorTable.get(color);
//		
//		if(xssfFillColor != null) {
//			return xssfFillColor;
//		}
//		colorTable.put(color, xssfFillColor);
//		return new XSSFColor(color.getRGB());
//	}
//	
//	/**
//	 * Utility to adapt ZSS 3.5 font into XSSFFont
//	 * @param 3.5 font
//	 * @return a XSSFFont Object
//	 */
//	protected XSSFFont toXSSFFont(NFont font) {
//
//		XSSFFont xssfFont = (XSSFFont) fontTable.get(font);
//		if(xssfFont != null) {
//			return xssfFont;
//		}
//		
//		xssfFont = (XSSFFont) workbook.createFont();
//		xssfFont.setBoldweight(toPOIBoldweight(font.getBoldweight()));
//		xssfFont.setStrikeout(font.isStrikeout());
//		xssfFont.setItalic(font.isItalic());
//		xssfFont.setColor(new XSSFColor(font.getColor().getRGB()));
//		xssfFont.setFontHeight(font.getHeightPoints());
//		xssfFont.setFontName(font.getName());
//		xssfFont.setTypeOffset(toPOITypeOffset(font.getTypeOffset()));
//		xssfFont.setUnderline(toPOIUnderline(font.getUnderline()));
//		
//		// put into table
//		fontTable.put(font, xssfFont);
//		
//		return xssfFont;
//	}

}
