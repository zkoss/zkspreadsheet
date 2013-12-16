/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/13 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import java.util.*;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NCellStyle.*;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NFont.*;


/**
 * Contains common importing behavior for both XLSX and XLS.
 * @author Hawk
 *
 */
abstract public class AbstractExcelImporter extends AbstractImporter {
	/** 
	 * <poi CellStyle index, {@link NCellStyle} object> 
	 * Keep track of imported style during importing to avoid creating duplicated style objects. 
	 */
	protected Map<Short, NCellStyle> importedStyle = new HashMap<Short, NCellStyle>();
	/**<poi Font index, {@link NFont} object> **/
	protected Map<Short, NFont> importedFont = new HashMap<Short, NFont>();
	protected NBook book;
	protected Workbook workbook;

	abstract protected int getMaxConfiguredColumn(Sheet poiSheet);
	
	/*
	 * import sheet scope content from POI Sheet.
	 */
	protected NSheet importPoiSheet(Sheet poiSheet) {
		NSheet sheet = book.createSheet(poiSheet.getSheetName());
		sheet.setDefaultRowHeight(XUtils.twipToPx(poiSheet.getDefaultRowHeight()));
		//TODO default char width reference XSSFBookImpl._defaultCharWidth = 7
		//TODO original width 64px, current is 72px
		//reference XUtils.getDefaultColumnWidthInPx()
		sheet.setDefaultColumnWidth(XUtils.defaultColumnWidthToPx(poiSheet.getDefaultColumnWidth(),8));
		for(Row poiRow : poiSheet) {
			importRow(sheet, poiRow);
		}
		
		//import columns
		for (int c=0 ; c <= getMaxConfiguredColumn(poiSheet) ; c++){
			//reference Spreadsheet.updateColWidth()
			sheet.getColumn(c).setWidth(XUtils.getWidthAny(poiSheet, c, 8));
			CellStyle columnStyle = poiSheet.getColumnStyle(c); 
			sheet.getColumn(c).setHidden(poiSheet.isColumnHidden(c));				
			if (columnStyle != null){
				sheet.getColumn(c).setCellStyle(importPoiCellStyle(columnStyle));
			}
		}
		
		return sheet;
	}
	
	protected NRow importRow(NSheet sheet, Row poiRow) {
		NRow row = sheet.getRow(poiRow.getRowNum());
		row.setHeight(XUtils.twipToPx(poiRow.getHeight()));
		CellStyle rowStyle = poiRow.getRowStyle();
		row.setHidden(poiRow.getZeroHeight());
		if (rowStyle != null){
			row.setCellStyle(importPoiCellStyle(rowStyle));
		}
		
		for(Cell poiCell : poiRow) {
			importPoiCell(sheet,poiRow.getRowNum(), poiCell);
		}
		
		return row;
	}
	
	protected NCell importPoiCell(NSheet sheet, int row, Cell poiCell){
		NCell cell = sheet.getCell(row,poiCell.getColumnIndex());
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
		cell.setCellStyle(importPoiCellStyle(poiCell.getCellStyle()));
		
		return cell;
	}
	
	/**
	 * Convert CellStyle into NCellStyle
	 * @param poiCellStyle
	 */
	protected NCellStyle importPoiCellStyle(CellStyle poiCellStyle) {
		NCellStyle cellStyle = null;
		if((cellStyle = importedStyle.get(poiCellStyle.getIndex())) ==null){
			cellStyle = book.createCellStyle(true);
			importedStyle.put(poiCellStyle.getIndex(), cellStyle);
			//reference XSSFDataFormat.getFormat()
			//FIXME currently we get data format converted upon locale
			cellStyle.setDataFormat(poiCellStyle.getDataFormatString());
			cellStyle.setWrapText(poiCellStyle.getWrapText());
			cellStyle.setLocked(poiCellStyle.getLocked());
			cellStyle.setAlignment(convertAlignment(poiCellStyle.getAlignment()));
			cellStyle.setVerticalAlignment(convertVerticalAlignment(poiCellStyle.getVerticalAlignment()));
			cellStyle.setFillColor(book.createColor(BookHelper.colorToBackgroundHTML(workbook, poiCellStyle.getFillForegroundColorColor())));

			cellStyle.setBorderLeft(convertBorderType(poiCellStyle.getBorderLeft()));
			cellStyle.setBorderTop(convertBorderType(poiCellStyle.getBorderTop()));
			cellStyle.setBorderRight(convertBorderType(poiCellStyle.getBorderRight()));
			cellStyle.setBorderBottom(convertBorderType(poiCellStyle.getBorderBottom()));

			cellStyle.setBorderLeftColor(book.createColor(BookHelper.colorToBorderHTML(workbook, poiCellStyle.getLeftBorderColorColor())));
			cellStyle.setBorderTopColor(book.createColor(BookHelper.colorToBorderHTML(workbook,poiCellStyle.getTopBorderColorColor())));
			cellStyle.setBorderRightColor(book.createColor(BookHelper.colorToBorderHTML(workbook,poiCellStyle.getRightBorderColorColor())));
			cellStyle.setBorderBottomColor(book.createColor(BookHelper.colorToBorderHTML(workbook,poiCellStyle.getBottomBorderColorColor())));
			cellStyle.setHidden(poiCellStyle.getHidden());
			cellStyle.setFillPattern(convertPoiFillPattern(poiCellStyle.getFillPattern()));
			//same style always use same font
			cellStyle.setFont(importFont(poiCellStyle));
		}
		
		return cellStyle;
	}	
	
	protected NFont importFont(CellStyle poiCellStyle) {
		NFont font = null;
		if (importedFont.containsKey(poiCellStyle.getFontIndex())){
			font = importedFont.get(poiCellStyle.getFontIndex());
		}else{
			Font poiFont = workbook.getFontAt(poiCellStyle.getFontIndex());
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
			font.setUnderline(convertUnderline(poiFont.getUnderline()));
			
			font.setHeightPoints(poiFont.getFontHeightInPoints());
			font.setTypeOffset(convertTypeOffset(poiFont.getTypeOffset()));
			font.setColor(book.createColor(BookHelper.getFontHTMLColor(workbook,poiFont)));
		}
		return font;
	}
	
	/*
	 * reference BookHelper.getFontCSSStyle()
	 */
	protected Underline convertUnderline(byte underline){
		switch(underline){
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
	
	protected TypeOffset convertTypeOffset(short typeOffset){
		switch(typeOffset){
		case Font.SS_SUB:
			return TypeOffset.SUB;
		case Font.SS_SUPER:
			return TypeOffset.SUPER;
		case Font.SS_NONE:
		default:
			return TypeOffset.NONE;
		}
	}
	
	protected BorderType convertBorderType(short poiBorder) {
		switch (poiBorder) {
			case CellStyle.BORDER_DASH_DOT:
				return BorderType.DASH_DOT;
			case CellStyle.BORDER_DASH_DOT_DOT:
				return BorderType.DASH_DOT_DOT;
			case CellStyle.BORDER_DASHED:
				return BorderType.DASHED;
			case CellStyle.BORDER_DOTTED:
				return BorderType.DOTTED;
			case CellStyle.BORDER_NONE:
				return BorderType.NONE;
			case CellStyle.BORDER_THIN:
			default:
				return BorderType.THIN; //unsupported border type
		}				
	}
	
	protected Alignment convertAlignment(short poiAlignment) {
		switch (poiAlignment) {
			case CellStyle.ALIGN_LEFT:
				return Alignment.LEFT;
			case CellStyle.ALIGN_RIGHT:
				return Alignment.RIGHT;
			case CellStyle.ALIGN_CENTER:
			case CellStyle.ALIGN_CENTER_SELECTION:
				return Alignment.CENTER;
			case CellStyle.ALIGN_FILL:
				return Alignment.FILL;
			case CellStyle.ALIGN_JUSTIFY:
				return Alignment.JUSTIFY;
			case CellStyle.ALIGN_GENERAL:
			default:	
				return Alignment.GENERAL;
		}
	}
	
	protected VerticalAlignment convertVerticalAlignment(short vAlignment) {
		switch (vAlignment) {
			case CellStyle.VERTICAL_TOP:
				return VerticalAlignment.TOP;
			case CellStyle.VERTICAL_CENTER:
				return VerticalAlignment.CENTER;
			case CellStyle.VERTICAL_JUSTIFY:
				return VerticalAlignment.JUSTIFY;
			case CellStyle.VERTICAL_BOTTOM:
			default:	
				return VerticalAlignment.BOTTOM;
		}
	}

	//TODO more error codes
	protected ErrorValue convertErrorCode(byte errorCellValue) {
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
	
	protected FillPattern convertPoiFillPattern(short poiFillPattern){
		switch(poiFillPattern){
			case CellStyle.SOLID_FOREGROUND:
				return FillPattern.SOLID_FOREGROUND;
			case CellStyle.FINE_DOTS:
				return FillPattern.FINE_DOTS;
			case CellStyle.ALT_BARS:
				return FillPattern.ALT_BARS;
			case CellStyle.SPARSE_DOTS:
				return FillPattern.SPARSE_DOTS;
			case CellStyle.THICK_HORZ_BANDS:
				return FillPattern.THICK_HORZ_BANDS;
			case CellStyle.THICK_VERT_BANDS:
				return FillPattern.THICK_VERT_BANDS;
			case CellStyle.THICK_BACKWARD_DIAG:
				return FillPattern.THICK_BACKWARD_DIAG;
			case CellStyle.THICK_FORWARD_DIAG:
				return FillPattern.THICK_FORWARD_DIAG;
			case CellStyle.BIG_SPOTS:
				return FillPattern.BIG_SPOTS;
			case CellStyle.BRICKS:
				return FillPattern.BRICKS;
			case CellStyle.THIN_HORZ_BANDS:
				return FillPattern.THIN_HORZ_BANDS;
			case CellStyle.THIN_VERT_BANDS:
				return FillPattern.THIN_VERT_BANDS;
			case CellStyle.THIN_BACKWARD_DIAG:
				return FillPattern.THIN_BACKWARD_DIAG;
			case CellStyle.THIN_FORWARD_DIAG:
				return FillPattern.THIN_FORWARD_DIAG;
			case CellStyle.SQUARES:
				return FillPattern.SQUARES;
			case CellStyle.DIAMONDS:
				return FillPattern.DIAMONDS;
			case CellStyle.LESS_DOTS:
				return FillPattern.LESS_DOTS;
			case CellStyle.LEAST_DOTS:
				return FillPattern.LEAST_DOTS;
			case CellStyle.NO_FILL:
			default:
			return FillPattern.NO_FILL;
		}
	}
}
