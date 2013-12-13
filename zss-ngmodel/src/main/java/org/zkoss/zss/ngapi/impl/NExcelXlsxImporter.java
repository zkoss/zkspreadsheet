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
package org.zkoss.zss.ngapi.impl;

import java.io.*;
import java.util.*;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;
import org.zkoss.zss.ngmodel.NFont.Underline;
/**
 * Convert Excel XLSX format file to a Spreadsheet {@link NBook} model including following information:
 * Book:
 * 		name
 * Sheet:
 * 		name, column width, row height, hidden row (column), row (column) style
 * Cell:
 * 		type, value, font with color and style, type offset(normal or subscript), background color, border
 * 
 * TODO use XLSX, XLS common interface (e.g. CellStyle instead of {@link XSSFCellStyle}) to get content first for that codes can be easily moved to parent class. 
 * @author Hawk
 * @since 3.5.0
 */
public class NExcelXlsxImporter extends AbstractExcelImporter{

	/** 
	 * <poi CellStyle index, {@link NCellStyle} object> 
	 * Keep track of imported style during importing to avoid creating duplicated style objects. 
	 */
	private Map<Short, NCellStyle> importedStyle = new HashMap<Short, NCellStyle>();
	/**<poi Font index, {@link NFont} object> **/
	private Map<Short, NFont> importedFont = new HashMap<Short, NFont>();
	// destination book model
	private NBook book;
	// source book model
	private XSSFWorkbook workbook;
	
	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {
		workbook = new XSSFWorkbook(is);
		book = NBooks.createBook(bookName);
		// import book scope content
		for(Sheet xssfSheet : workbook) {
			importPoiSheet(xssfSheet);
		}
		return book;
	}


	/*
	 * import sheet scope content from XSSFSheet.
	 */
	private NSheet importPoiSheet(Sheet poiSheet) {
		NSheet sheet = book.createSheet(poiSheet.getSheetName());
		sheet.setDefaultRowHeight(XUtils.twipToPx(poiSheet.getDefaultRowHeight()));
		//TODO default char width reference XSSFBookImpl._defaultCharWidth = 7
		//TODO original width 64px, current is 72px
		//reference XUtils.getDefaultColumnWidthInPx()
		sheet.setDefaultColumnWidth(XUtils.defaultColumnWidthToPx(poiSheet.getDefaultColumnWidth(),8));
		int maxColumnIndex = 0;
		for(Row poiRow : poiSheet) {
			importRow(sheet, poiRow);
			if (poiRow.getLastCellNum() > maxColumnIndex){
				maxColumnIndex = poiRow.getLastCellNum(); 
			}
		}
		
		//import columns
		for (int c=0 ; c<=maxColumnIndex ; c++){
			//reference Spreadsheet.updateColWidth()
			sheet.getColumn(c).setWidth(XUtils.getWidthAny(poiSheet, c, 8));
			if (poiSheet.getColumnStyle(c) != null){
				sheet.getColumn(c).setCellStyle(importXSSFCellStyle((XSSFCellStyle)poiSheet.getColumnStyle(c)));
			}
		}
		
		return sheet;
	}

	private NRow importRow(NSheet sheet, Row poiRow) {
		NRow row = sheet.getRow(poiRow.getRowNum());
		row.setHeight(XUtils.twipToPx(poiRow.getHeight()));
		if (poiRow.getRowStyle()!= null){
			row.setCellStyle(importXSSFCellStyle((XSSFCellStyle)poiRow.getRowStyle()));
		}
		
		for(Cell poiCell : poiRow) {
			importPoiCell(row, poiCell);
		}
		
		return row;
	}
	
	private NCell importPoiCell(NRow row, Cell poiCell){
		NCell cell = row.getCell(poiCell.getColumnIndex());
		switch (poiCell.getCellType()){
			case Cell.CELL_TYPE_NUMERIC:
				cell.setNumberValue(poiCell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				cell.setStringValue(poiCell.getStringCellValue());
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cell.setBooleanValue(poiCell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_FORMULA:
				cell.setFormulaValue(poiCell.getCellFormula());
				break;
			case Cell.CELL_TYPE_ERROR:
				cell.setErrorValue(convertErrorCode(poiCell.getErrorCellValue()));
				break;
			case Cell.CELL_TYPE_BLANK:
				//do nothing because spreadsheet model auto creates blank cells
				break;
			default:
				//TODO log "ignore a cell with unknown.
		}
		cell.setCellStyle(importXSSFCellStyle((XSSFCellStyle)poiCell.getCellStyle()));
		
		return cell;
	}
	
	/**
	 * Convert CellStyle into NCellStyle
	 * @param xssfCellStyle
	 */
	private NCellStyle importXSSFCellStyle(XSSFCellStyle xssfCellStyle) {
		NCellStyle cellStyle = null;
		if((cellStyle = importedStyle.get(xssfCellStyle.getIndex())) ==null){
			cellStyle = book.createCellStyle(true);
			importedStyle.put(xssfCellStyle.getIndex(), cellStyle);
			//FIXME this will get converted data format pattern
			cellStyle.setDataFormat(xssfCellStyle.getDataFormatString());
			cellStyle.setWrapText(xssfCellStyle.getWrapText());
			cellStyle.setLocked(xssfCellStyle.getLocked());
			cellStyle.setAlignment(convertAlignment(xssfCellStyle.getAlignmentEnum()));
			cellStyle.setVerticalAlignment(convertVerticalAlignment(xssfCellStyle.getVerticalAlignmentEnum()));
			cellStyle.setFillColor(convertPoiColor(xssfCellStyle.getFillForegroundColorColor(),"#000000"));

			cellStyle.setBorderLeft(convertBorderType(xssfCellStyle.getBorderLeftEnum()));
			cellStyle.setBorderTop(convertBorderType(xssfCellStyle.getBorderTopEnum()));
			cellStyle.setBorderRight(convertBorderType(xssfCellStyle.getBorderRightEnum()));
			cellStyle.setBorderBottom(convertBorderType(xssfCellStyle.getBorderBottomEnum()));

			cellStyle.setBorderLeftColor(convertPoiColor(xssfCellStyle.getLeftBorderColorColor(), "#FFFFFF"));
			cellStyle.setBorderTopColor(convertPoiColor(xssfCellStyle.getTopBorderColorColor(), "#FFFFFF"));
			cellStyle.setBorderRightColor(convertPoiColor(xssfCellStyle.getRightBorderColorColor(), "#FFFFFF"));
			cellStyle.setBorderBottomColor(convertPoiColor(xssfCellStyle.getBottomBorderColorColor(), "#FFFFFF"));
			cellStyle.setHidden(xssfCellStyle.getHidden());
			/*
//			nCellStyle.setFillPattern(fillPattern);
			 */
			//same style always use same font
			cellStyle.setFont(importFont(xssfCellStyle));
		}
		
		return cellStyle;
		
	}


	private NFont importFont(XSSFCellStyle xssfCellStyle) {
		NFont font = null;
		if (importedFont.containsKey(xssfCellStyle.getFontIndex())){
			font = importedFont.get(xssfCellStyle.getFontIndex());
		}else{
			Font poiFont = workbook.getFontAt(xssfCellStyle.getFontIndex());
			font = book.createFont(true);
			//font
			font.setName(poiFont.getFontName());
			if (poiFont.getBoldweight()==Font.BOLDWEIGHT_BOLD){
				font.setBoldweight(NFont.Boldweight.BOLD);
			}else{
				font.setBoldweight(NFont.Boldweight.NORMAL);
			}
			font.setItalic(poiFont.getItalic());
			font.setStrikeout(poiFont.getStrikeout());
			font.setUnderline(convertUnderline(poiFont));
			
			font.setHeightPoints(poiFont.getFontHeightInPoints());
			font.setTypeOffset(convertTypeOffset(poiFont));
			font.setColor(convertPoiColor(((XSSFFont)poiFont).getXSSFColor(),"#FFFFFF"));
		}
		return font;
	}

	//FIXME how to handle AUTO_COLOR
	//BookHelper already have 
	private NColor convertPoiColor(Color color, String autoColor){
		String htmlColor = BookHelper.colorToHTML(workbook, color);
		if ("AUTO_COLOR".equals(htmlColor)){
			htmlColor = autoColor;
		}
		return book.createColor(htmlColor);
	}
	/*
	 * reference BookHelper.getFontCSSStyle()
	 */
	private Underline convertUnderline(Font poiFont){
		switch(poiFont.getUnderline()){
		case Font.U_SINGLE:
			return NFont.Underline.SINGLE;
		case Font.U_DOUBLE:
			return NFont.Underline.DOUBLE;
		case Font.U_SINGLE_ACCOUNTING:
			return NFont.Underline.SINGLE_ACCOUNTING;
		case Font.U_DOUBLE_ACCOUNTING:
			return NFont.Underline.DOUBLE_ACCOUNTING;
		case Font.U_NONE:
		default:
			return NFont.Underline.NONE;
		}
	}
	
	private TypeOffset convertTypeOffset(Font poiFont){
		switch(poiFont.getTypeOffset()){
		case Font.SS_SUB:
			return TypeOffset.SUB;
		case Font.SS_SUPER:
			return TypeOffset.SUPER;
		case Font.SS_NONE:
		default:
			return TypeOffset.NONE;
		}
	}
	
	private BorderType convertBorderType(BorderStyle poiBorder) {
		switch (poiBorder) {
			case DASH_DOT:
				return BorderType.DASH_DOT;
			case DASH_DOT_DOT:
				return BorderType.DASH_DOT_DOT;
			case DASHED:
				return BorderType.DASHED;
			case DOTTED:
				return BorderType.DOTTED;
			case NONE:
				return BorderType.NONE;
			case THIN:
			default:
				return BorderType.THIN; //unsupported border type
		}				
	}
	
	private Alignment convertAlignment(HorizontalAlignment poiAlignment) {
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
	
	private VerticalAlignment convertVerticalAlignment(org.zkoss.poi.ss.usermodel.VerticalAlignment vAlignment) {
		switch (vAlignment) {
			case TOP:
				return VerticalAlignment.TOP;
			case CENTER:
				return VerticalAlignment.CENTER;
			case JUSTIFY:
				return VerticalAlignment.JUSTIFY;
			case BOTTOM:
			default:	
				return VerticalAlignment.BOTTOM;
		}
	}

	//TODO more error codes
	private ErrorValue convertErrorCode(byte errorCellValue) {
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
}
 