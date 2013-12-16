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

import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.xssf.usermodel.*;
import org.zkoss.zss.ngmodel.*;
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
		for(Sheet poiSheet : workbook) {
			importPoiSheet(poiSheet);
		}
		return book;
	}


	/*
	 * import sheet scope content from POI Sheet.
	 */
	private NSheet importPoiSheet(Sheet poiSheet) {
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
	
	/**
	 * get last column index that is configured, e.g. 16383
	 * reference BookHelper.getMaxConfiguredColumn()
	 * @param poiSheet
	 * @return
	 */
	private int getMaxConfiguredColumn(Sheet poiSheet) {
		int max = -1;
		CTWorksheet worksheet = ((XSSFSheet)poiSheet).getCTWorksheet();
		if(worksheet.sizeOfColsArray()<=0){
			return max;
		}
		CTCols colsArray = worksheet.getColsArray(0);
		for (int i = 0; i < colsArray.sizeOfColArray(); i++) {
			CTCol colArray = colsArray.getColArray(i);
			max = Math.max(max, (int) colArray.getMax()-1);//in CT col it is 1base, -1 to 0base.
		}
		return max;
	}

	private NRow importRow(NSheet sheet, Row poiRow) {
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
	
	private NCell importPoiCell(NSheet sheet, int row, Cell poiCell){
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
	private NCellStyle importPoiCellStyle(CellStyle poiCellStyle) {
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


	private NFont importFont(CellStyle poiCellStyle) {
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
			font.setColor(book.createColor(converFontColor(poiFont)));
		}
		return font;
	}
	
	private String converFontColor(Font font){
		XSSFFont f = (XSSFFont) font;
		XSSFColor color = f.getXSSFColor();
		return BookHelper.colorToHTML(workbook, color);
	}
}
 