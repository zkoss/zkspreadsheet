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
import org.zkoss.zss.ngmodel.impl.BookImpl;
/**
 * Convert Excel XLSX format file to a Spreadsheet {@link NBook} model including following information:
 * Book:
 * 		name
 * Sheet:
 * 		name, column width, row height, hidden row (column), row (column) style
 * Cell:
 * 		type, value, font with color and style, type offset(normal or subscript), background color, border 
 * @author Hawk
 * @since 3.5.0
 */
public class NExcelXlsxImporter extends AbstractImporter{

	//<XSSF style index, NCellStyle object>
	private Map<Short, NCellStyle> importedStyle = new HashMap<Short, NCellStyle>();
	private Map<Short, NFont> importedFont = new HashMap<Short, NFont>();
	private NBook book;
	private XSSFWorkbook workbook;
	
	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {
		workbook = new XSSFWorkbook(is);
		book = new BookImpl(bookName);
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
			sheet.getColumn(c).setCellStyle(importXSSFCellStyle((XSSFCellStyle)poiSheet.getColumnStyle(c)));
		}
	}

	private void importRow(NSheet sheet, Row poiRow) {
		NRow row = sheet.getRow(poiRow.getRowNum());
		row.setHeight(XUtils.twipToPx(poiRow.getHeight()));
		if (poiRow.getRowStyle()!= null){
			row.setCellStyle(importXSSFCellStyle((XSSFCellStyle)poiRow.getRowStyle()));
		}
		
		for(Cell poiCell : poiRow) { // Go through each cell
			NCell cell = importPoiCell(sheet, poiCell);
			//same style always use same font
			if(!importedStyle.containsKey(poiCell.getCellStyle().getIndex())) {
				cell.setCellStyle(importXSSFCellStyle((XSSFCellStyle)poiCell.getCellStyle()));
			}
		}
	}
	
	/**
	 * copy XSSFCellStyle attributes into nCellStyle
	 * @param nCellStyle
	 * @param xssfCellStyle
	 */
	private NCellStyle importXSSFCellStyle(XSSFCellStyle xssfCellStyle) {
		NCellStyle cellStyle;
			cellStyle = book.createCellStyle(true);
			importedStyle.put(xssfCellStyle.getIndex(), cellStyle);
			
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
			/*
			cellStyle.setHidden(xssfCellStyle.getHidden());
//			nCellStyle.setFillPattern(fillPattern);
			 */
			
			cellStyle.setFont(importFont(xssfCellStyle));
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

}
 