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

import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.ErrorConstants;
import org.zkoss.poi.ss.usermodel.HorizontalAlignment;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Workbook;
import org.zkoss.poi.xssf.usermodel.XSSFCell;
import org.zkoss.poi.xssf.usermodel.XSSFCellStyle;
import org.zkoss.poi.xssf.usermodel.XSSFFont;
import org.zkoss.poi.xssf.usermodel.XSSFRow;
import org.zkoss.poi.xssf.usermodel.XSSFSheet;
import org.zkoss.poi.xssf.usermodel.XSSFWorkbook;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.NFont;
import org.zkoss.zss.ngmodel.NRow;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.impl.BookImpl;
import org.zkoss.zss.ngmodel.impl.ColorImpl;
import org.zkoss.zss.ngmodel.impl.FontImpl;
/**
 * Convert Excel XLSX format file to Spreadsheet book model including following information:
 * Book:
 * 		name
 * Sheet:
 * 		name
 * Cell:
 * 		type, value, font with color and style
 * @author dennis
 * @since 3.5.0
 */
public class NExcelXlsxImporter extends AbstractImporter{


	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {

		Workbook workbook = new XSSFWorkbook(is);
		NBook book = new BookImpl(bookName);
		importXSSFBook(book, (XSSFWorkbook) workbook);
		
		return book;
	}

	/**
	 * Copy XSSFBook attributes into NBook
	 * 
	 * @param nBook
	 * @param xssfBook
	 */
	private void importXSSFBook(NBook nBook, XSSFWorkbook xssfBook) {
		
		// FIXME: Should I handle style in the beginning?
		/**
		 * Handle book style table
		 */
		//StylesTable styleTable = xssfBook.getStylesSource();		
		/*
		 * Copy XSSFFont into NFont.
		 * Create Font into table for each Font.
		 */		
//		for(XSSFFont xssfFont : styleTable.getFonts()) {
//			
//			NFont nFont = new FontImpl();
//			
//			NColor nColor = new ColorImpl(xssfFont.getXSSFColor().getRgb());
//			// FIXME
//			// Should I create color?
//			// nBook.createColor();
//			
//			nFont.setColor(nColor);
//			nFont.setBoldweight(NFont.Boldweight.values()[xssfFont.getBoldweight()]);
//			nFont.setHeight(xssfFont.getFontHeight());
//			nFont.setItalic(xssfFont.getItalic());
//			nFont.setName(xssfFont.getFontName());
//			nFont.setStrikeout(xssfFont.getStrikeout());
//			nFont.setTypeOffset(NFont.TypeOffset.values()[xssfFont.getTypeOffset()]);
//			nFont.setUnderline(NFont.Underline.values()[xssfFont.getUnderline()]);
//			
//			bookAdv.createFont(nFont ,true);
//		}
		
		// Go through all sheet in book
		for(XSSFSheet xssfSheet : xssfBook) {
			NSheet nSheet = nBook.createSheet(xssfSheet.getSheetName());
			importXSSFSheet(nSheet, xssfSheet);
		}
	}

	/**
	 * Copy XSSFSheet attributes into NSheet
	 * @param nSheet
	 * @param xssfSheet
	 */
	private void importXSSFSheet(NSheet nSheet, XSSFSheet xssfSheet) {
		
		for(Row row : xssfSheet) { // Go through each row
			
			XSSFRow xssfRow = (XSSFRow) row;
			NRow nRow = nSheet.getRow(xssfRow.getRowNum());
			
			for(Cell cell : xssfRow) { // Go through each cell
				
				XSSFCell xssfCell = (XSSFCell) cell;
				NCell nCell = importXSSFCell(nSheet, xssfCell);
				nCell.setCellStyle(importXSSFCellStyle(nCell, xssfCell.getCellStyle()));
				
				// TODO: copy hyper link
				// nCell.getHyperlink();
			}
			
		}
		
	}
	
	private NCell importXSSFCell(NSheet sheet, XSSFCell xssfCell){
		NCell cell = sheet.getCell(xssfCell.getRowIndex(), xssfCell.getColumnIndex());
		switch (xssfCell.getCellType()){
			case Cell.CELL_TYPE_NUMERIC:
				cell.setNumberValue(xssfCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setStringValue(xssfCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cell.setBooleanValue(xssfCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				cell.setFormulaValue(xssfCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_ERROR:
				cell.setErrorValue(convertErrorCode(xssfCell.getErrorCellValue()));
				break;
			case Cell.CELL_TYPE_BLANK:
				//do nothing because spreadsheet model auto creates blank cells
				break;
			default:
				//TODO log "ignore a cell with unknown.
		}
		return cell;
	}
	
	private ErrorValue convertErrorCode(byte errorCellValue){
		switch (errorCellValue){
			case ErrorConstants.ERROR_NAME:
				return new ErrorValue(ErrorValue.INVALID_NAME);
			case ErrorConstants.ERROR_VALUE:
				return new ErrorValue(ErrorValue.INVALID_VALUE);
			default:
				//TODO log it
				return new ErrorValue(ErrorValue.INVALID_NAME);
		}
		
	}
	/**
	 * copy XSSFCellStyle attributes into nCellStyle
	 * @param nCellStyle
	 * @param xssfCellStyle
	 */
	private NCellStyle importXSSFCellStyle(NCell nCell, XSSFCellStyle xssfCellStyle) {
		
		NCellStyle nCellStyle = null;
		
		if(!isStyleExist(xssfCellStyle.getIndex())) {
			
			nCellStyle = nCell.getCellStyle();
		
			XSSFFont xssfFont = xssfCellStyle.getFont();
			//FIXME should create font
			NFont nFont = nCell.getSheet().getBook().createFont(true);
			
			NColor nColor = new ColorImpl(xssfFont.getXSSFColor().getRgb());
			
			// FIXME
			// Should I create color?
			// nBook.createColor();
			
			nFont.setColor(nColor);
			// FIXME: ENUM
			nFont.setBoldweight(NFont.Boldweight.values()[xssfFont.getBoldweight()]);
			nFont.setHeightPoints(xssfFont.getFontHeightInPoints());
			nFont.setItalic(xssfFont.getItalic());
			nFont.setName(xssfFont.getFontName());
			nFont.setStrikeout(xssfFont.getStrikeout());
			// FIXME: ENUM
			nFont.setTypeOffset(NFont.TypeOffset.values()[xssfFont.getTypeOffset()]);
			// FIXME: ENUM
			nFont.setUnderline(NFont.Underline.values()[xssfFont.getUnderline()]);
			
			// FIXME
			// Should I create font?
			// nBook.createFont();
			
			nCellStyle.setLocked(xssfCellStyle.getLocked());
			
			
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
			nCellStyle.setFont(nFont);
			nCellStyle.setHidden(xssfCellStyle.getHidden());
//			nCellStyle.setVerticalAlignment(verticalAlignment);
			nCellStyle.setWrapText(xssfCellStyle.getWrapText());

			
		}
		
		return nCellStyle;
		
	}
	
	/**
	 * TODO
	 * Utility
	 * does style exist ID in the POI style table
	 * @return
	 */
	private boolean isStyleExist(short styleID) {
		return false;
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
 