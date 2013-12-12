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

import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.HorizontalAlignment;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.xssf.usermodel.XSSFCellStyle;
import org.zkoss.poi.xssf.usermodel.XSSFFont;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBooks;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;
import org.zkoss.zss.ngmodel.NFont.Underline;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NSheet;
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

	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {

		XSSFWorkbook workbook = new XSSFWorkbook(is);
		NBook book = NBooks.createBook(bookName);
		// Go through all sheet in book
		for(XSSFSheet xssfSheet : workbook) {
			NSheet nSheet = book.createSheet(xssfSheet.getSheetName());
			importXSSFSheet(nSheet, xssfSheet);
		}
		return book;
	}


	/**
	 * Copy XSSFSheet attributes into NSheet
	 * @param sheet
	 * @param xssfSheet
	 */
	private void importXSSFSheet(NSheet sheet, XSSFSheet xssfSheet) {
		
		for(Row poiRow : xssfSheet) { // Go through each row
			
			NRow row = sheet.getRow(poiRow.getRowNum());
			row.setHeight(XUtils.twipToPx(poiRow.getHeight()));
			for(Cell poiCell : poiRow) { // Go through each cell
				
				NCell cell = importPoiCell(sheet, poiCell);
				cell.setCellStyle(importXSSFCellStyle(cell, (XSSFCellStyle)poiCell.getCellStyle()));
				
				// TODO: copy hyper link
				// nCell.getHyperlink();
			}
			
		}
		
	}
	
	/**
	 * copy XSSFCellStyle attributes into nCellStyle
	 * @param nCellStyle
	 * @param xssfCellStyle
	 */
	private NCellStyle importXSSFCellStyle(NCell cell, XSSFCellStyle xssfCellStyle) {
		NCellStyle cellStyle = retrieveCellStyle(cell, xssfCellStyle);
		
		//font
		XSSFFont xssfFont = xssfCellStyle.getFont();
		NFont font = cell.getSheet().getBook().createFont(true);
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
	 * Retrieve corresponding CellStyle if same XSSFCellStyle is imported before.
	 * Otherwise, create a new one. 
	 */
	private NCellStyle retrieveCellStyle(NCell cell, XSSFCellStyle xssfCellStyle) {
		NCellStyle cellStyle;
		if(importedStyle.containsKey(xssfCellStyle.getIndex())) {
			cellStyle = importedStyle.get(xssfCellStyle.getIndex());
		}else{
			cellStyle = cell.getSheet().getBook().createCellStyle(true);
			cell.setCellStyle(cellStyle);
			importedStyle.put(xssfCellStyle.getIndex(), cellStyle);
		}
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
 