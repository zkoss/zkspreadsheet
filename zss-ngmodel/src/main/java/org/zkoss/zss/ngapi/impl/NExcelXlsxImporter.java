/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl;

import java.io.*;
import java.util.*;

import org.zkoss.poi.hssf.converter.ExcelToHtmlUtils;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;
import org.zkoss.zss.ngmodel.NFont.Underline;
import org.zkoss.zss.ngmodel.impl.BookImpl;
/**
 * Convert Excel XLSX format file to a Spreadsheet {@link NBook} model including following information:
 * Book:
 * 		name
 * Sheet:
 * 		name, column width, row height, hidden row (column), 
 * Cell:
 * 		type, value, font with color and style, type offset(normal or subscript), background color, border 
 * @author dennis
 * @since 3.5.0
 */
public class NExcelXlsxImporter extends AbstractImporter{

	//<XSSF style index, NCellStyle object>
	private Map<Short, NCellStyle> importedStyle = new HashMap<Short, NCellStyle>();
	private Map<Short, NFont> importedFont = new HashMap<Short, NFont>();

	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook(is);
		NBook book = new BookImpl(bookName);
		// import book scope content
		for(XSSFSheet xssfSheet : workbook) {
			importXSSFSheet(book, xssfSheet);
		}
		return book;
	}


	/*
	 * import sheet scope content from XSSFSheet.
	 */
	private void importXSSFSheet(NBook book, Sheet poiSheet) {
		NSheet sheet = book.createSheet(poiSheet.getSheetName());
		sheet.setDefaultRowHeight(XUtils.twipToPx(poiSheet.getDefaultRowHeight()));
		//TODO default char width reference XSSFBookImpl._defaultCharWidth
		sheet.setDefaultColumnWidth(XUtils.defaultColumnWidthToPx(poiSheet.getDefaultColumnWidth(),8));
		int maxColumnIndex = 0;
		for(Row poiRow : poiSheet) { // Go through each row
			importRow(sheet, poiRow);
			if (poiRow.getLastCellNum() > maxColumnIndex){
				maxColumnIndex = poiRow.getLastCellNum(); 
			}
		}
		
		for (int c=0 ; c<=maxColumnIndex ; c++){
//			sheet.getColumn(c).setWidth(ExcelToHtmlUtils.getColumnWidthInPx(poiSheet.getColumnWidth(c)));
		}
	}


	private void importRow(NSheet sheet, Row poiRow) {
		NRow row = sheet.getRow(poiRow.getRowNum());
		row.setHeight(XUtils.twipToPx(poiRow.getHeight()));
//		if (poiRow.getRowStyle()!= null)
//		row.setCellStyle()
		for(Cell poiCell : poiRow) { // Go through each cell
			
			NCell cell = importPoiCell(sheet, poiCell);
			//same style always use same font
			if(!importedStyle.containsKey(poiCell.getCellStyle().getIndex())) {
				cell.setCellStyle(importXSSFCellStyle(cell, (XSSFCellStyle)poiCell.getCellStyle()));
			}
			// TODO: copy hyper link
			// nCell.getHyperlink();
		}
	}
	
	/**
	 * copy XSSFCellStyle attributes into nCellStyle
	 * @param nCellStyle
	 * @param xssfCellStyle
	 */
	private NCellStyle importXSSFCellStyle(NCell cell, XSSFCellStyle xssfCellStyle) {
		NCellStyle cellStyle;
			cellStyle = cell.getSheet().getBook().createCellStyle(true);
			cell.setCellStyle(cellStyle);
			importedStyle.put(xssfCellStyle.getIndex(), cellStyle);
		
			NFont font = null;
			if (importedFont.containsKey(xssfCellStyle.getFontIndex())){
				font = importedFont.get(xssfCellStyle.getFontIndex());
			}else{
				XSSFFont xssfFont = xssfCellStyle.getFont();
				font = cell.getSheet().getBook().createFont(true);
				//font
				font.setName(xssfFont.getFontName());
				if (xssfFont.getBold()){
					font.setBoldweight(NFont.Boldweight.BOLD);
				}else{
					font.setBoldweight(NFont.Boldweight.NORMAL);
				}
				font.setItalic(xssfFont.getItalic());
				font.setStrikeout(xssfFont.getStrikeout());
				font.setUnderline(convertUnderline(xssfFont));
				
				font.setHeightPoints(xssfFont.getFontHeightInPoints());
				font.setTypeOffset(convertTypeOffset(xssfFont));
				// nBook.createFont();
				//			NColor color = cell.getSheet().getBook().createColor(xssfFont.getXSSFColor().getARGBHex());
				//			font.setColor(color);
			}
			cellStyle.setFont(font);
			// FIXME
			
//		cellStyle.setLocked(xssfCellStyle.getLocked());
			
			/*
			nCellStyle.setAlignment(poiToNGAlignment(xssfCellStyle.getAlignmentEnum()));
			nCellStyle.setBorderBottom(poiToBorderType(xssfCellStyle.getBorderBottomEnum()));
			nCellStyle.setBorderBottomColor(new ColorImpl(xssfCellStyle.getBottomBorderColorColor().getRgb()));
			nCellStyle.setBorderLeft(poiToBorderType(xssfCellStyle.getBorderLeftEnum()));
			nCellStyle.setBorderLeftColor(new ColorImpl(xssfCellStyle.getLeftBorderColorColor().getRgb()));
			nCellStyle.setBorderTop(poiToBorderType(xssfCellStyle.getBorderTopEnum()));
			nCellStyle.setBorderTopColor(new ColorImpl(xssfCellStyle.getTopBorderColorColor().getRgb()));
			nCellStyle.setBorderRight(poiToBorderType(xssfCellStyle.getBorderRightEnum()));
			nCellStyle.setBorderRightColor(new ColorImpl(xssfCellStyle.getRightBorderColorColor().getRgb()));
			nCellStyle.setDataFormat(xssfCellStyle.getDataFormatString());
			//nCellStyle.setFillColor(fillColor);
//			nCellStyle.setFillPattern(fillPattern);
			nCellStyle.setFont(font);
			nCellStyle.setHidden(xssfCellStyle.getHidden());
//			nCellStyle.setVerticalAlignment(verticalAlignment);
			nCellStyle.setWrapText(xssfCellStyle.getWrapText());
			 */
			
		return cellStyle;
		
	}


	/*
	 * reference BookHelper.getFontCSSStyle()
	 */
	private Underline convertUnderline(XSSFFont xssfFont){
		switch(xssfFont.getUnderline()){
		case XSSFFont.U_SINGLE:
			return NFont.Underline.SINGLE;
		case XSSFFont.U_DOUBLE:
			return NFont.Underline.DOUBLE;
		case XSSFFont.U_SINGLE_ACCOUNTING:
			return NFont.Underline.SINGLE_ACCOUNTING;
		case XSSFFont.U_DOUBLE_ACCOUNTING:
			return NFont.Underline.DOUBLE_ACCOUNTING;
		case XSSFFont.U_NONE:
		default:
			return NFont.Underline.NONE;
		}
	}
	
	private TypeOffset convertTypeOffset(XSSFFont xssfFont){
		switch(xssfFont.getTypeOffset()){
		case Font.SS_SUB:
			return TypeOffset.SUB;
		case Font.SS_SUPER:
			return TypeOffset.SUPER;
		case Font.SS_NONE:
		default:
			return TypeOffset.NONE;
		}
	}
	
	private BorderType poiToBorderType(BorderStyle poiBorder) {
		switch (poiBorder) {
			case THIN:
				return BorderType.THIN;
			case DASH_DOT:
				return BorderType.DASH_DOT;
			case DASH_DOT_DOT:
				return BorderType.DASH_DOT_DOT;
			case DASHED:
				return BorderType.DASHED;
			case DOTTED:
				return BorderType.DOTTED;
			default:
				return BorderType.NONE;
		}				
	}
	
	private Alignment poiToNGAlignment(HorizontalAlignment poiAlignment) {
		switch (poiAlignment) {
			case LEFT:
				return Alignment.LEFT;
			case RIGHT:
				return Alignment.RIGHT;
			case CENTER:
				return Alignment.CENTER;
			case FILL:
				return Alignment.FILL;
			case JUSTIFY:
				return Alignment.JUSTIFY;
			default:	
				return Alignment.GENERAL;
		}
	}

}
 