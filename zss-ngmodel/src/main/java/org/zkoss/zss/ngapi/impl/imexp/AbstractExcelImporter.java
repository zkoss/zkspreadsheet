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

import java.io.*;
import java.util.*;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFCellStyle;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NAutoFilter.FilterOp;
import org.zkoss.zss.ngmodel.NAutoFilter.NFilterColumn;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.FillPattern;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;
import org.zkoss.zss.ngmodel.NFont.Underline;
import org.zkoss.zss.ngmodel.NPicture.Format;

/**
 * Contains common importing behavior for both XLSX and XLS. Spreadsheet {@link NBook} model including following information:
 * Book:
 * 		name
 * Sheet:
 * 		name, (default) column width, (default) row height, hidden row (column), row (column) style, freeze, merge, protection, named range
 * , gridline display
 * Cell:
 * 		type, value, font with color and style, type offset(normal or subscript), background color, border's type and color
 * , data format, alignment, wrap, locked, fill pattern
 * 
 * We use XLSX, XLS common interface (e.g. CellStyle instead of {@link XSSFCellStyle}) to get content first for that codes can be easily moved to parent class. 
 * @author Hawk
 * @since 3.5.0
 */
abstract public class AbstractExcelImporter extends AbstractImporter {
	/**
	 * Office Open XML Part 4: Markup Language Reference
	 * 3.3.1.12  col (Column Width & Formatting)
	 * The character width 7 is based on Calibri 11.
	 * We can get correct column width under Excel 2007, but incorrect column width in 2010
	 */
	public static final int CHRACTER_WIDTH = 7;
	/** 
	 * <poi CellStyle index, {@link NCellStyle} object> 
	 * Keep track of imported style during importing to avoid creating duplicated style objects. 
	 */
	protected Map<Short, NCellStyle> importedStyle = new HashMap<Short, NCellStyle>();
	/** <poi Font index, {@link NFont} object> **/
	protected Map<Short, NFont> importedFont = new HashMap<Short, NFont>();
	/** target book model */
	protected NBook book;
	/** source POI book */
	protected Workbook workbook;

	/**
	 * Import the model according to reversed dependency order among model objects: 
	 * 		book, sheet, defined name, cells, chart, pictures, validation.
	 */
	@Override
	public NBook imports(InputStream is, String bookName) throws IOException {

		workbook = createPoiBook(is);
		book = NBooks.createBook(bookName);
		
		NBookSeries bookSeries = book.getBookSeries();
		boolean isCacheClean = bookSeries.isAutoFormulaCacheClean();
		try{
			bookSeries.setAutoFormulaCacheClean(false);//disable it to avoid unnecessary clean up during importing
			
			importExternalBookLinks();
			int numberOfSheet = workbook.getNumberOfSheets();
			for (int i = 0 ; i < numberOfSheet; i++){
				importSheet(workbook.getSheetAt(i));
			}
			importNamedRange();
			for (int i = 0 ; i < numberOfSheet; i++){
				NSheet sheet = book.getSheet(i);
				Sheet poiSheet = workbook.getSheetAt(i);
				
				for(Row poiRow : poiSheet) {
					importRow(poiRow, sheet);
				}
				importColumn(poiSheet, sheet);
				importMergedRegions(poiSheet, sheet);
				importDrawings(poiSheet, sheet);
				importValidation(poiSheet, sheet);
				importAutoFilter(poiSheet, sheet);
			}
		}finally{
			book.getBookSeries().setAutoFormulaCacheClean(isCacheClean);
		}
		
		return book;
	}
	
	abstract protected Workbook createPoiBook(InputStream is) throws IOException;
	
	/**
	 * When a column is hidden with default width, we don't import the width for it's 0.
	 * @param poiSheet
	 * @param sheet
	 */
	abstract protected void importColumn(Sheet poiSheet, NSheet sheet);
	
	/**
	 * If in same column:
	 * anchorWidthInFirstColumn + anchor width in inter-columns + anchorWidthInLastColumn (dx2)
	 * no in same column:
	 * anchorWidthInLastColumn - offsetInFirstColumn (dx1)
	 * 
	 */
	abstract protected int getAnchorWidthInPx(ClientAnchor anchor, Sheet poiSheet);
	abstract protected int getAnchorHeightInPx(ClientAnchor anchor, Sheet poiSheet);
	
	/**
	 * Name should be created after sheets created.
	 * A special defined name, _xlnm._FilterDatabase (xlsx) or _FilterDatabase (xls), stores the selected cells for auto-filter 
	 */
	protected void importNamedRange(){
		for (int i=0 ; i<workbook.getNumberOfNames() ; i++){
			Name definedName = workbook.getNameAt(i);
			if (definedName.isFunctionName() //ignore defined name of functions, they are macro functions that we don't support 
				|| definedName.getRefersToFormula()==null){ //ignore defined name with null formula, don't know when will have this case
				
				continue;
			}
			
			NName namedRange = null;
			if (definedName.getSheetIndex() == -1){//workbook scope
				namedRange = book.createName(definedName.getNameName());
			}else{
				namedRange = book.createName(definedName.getNameName(), definedName.getSheetName());
			}
			namedRange.setRefersToFormula(definedName.getRefersToFormula());
		}
	}
	
	/**
	 * Excel uses external book links to map external book index and name.
	 * The formula contains full external book name or index only (e.g [book2.xlsx] or [1]).
	 * We needs such table for parsing and evaluating formula when necessary. 
	 */
	abstract protected void importExternalBookLinks();
	
	/*
	 * import sheet scope content from POI Sheet.
	 */
	protected NSheet importSheet(Sheet poiSheet) {
		NSheet sheet = book.createSheet(poiSheet.getSheetName());
		sheet.setDefaultRowHeight(UnitUtil.twipToPx(poiSheet.getDefaultRowHeight()));
		//reference XUtils.getDefaultColumnWidthInPx()
		int defaultWidth = UnitUtil.defaultColumnWidthToPx(poiSheet.getDefaultColumnWidth(), CHRACTER_WIDTH);
		sheet.setDefaultColumnWidth(defaultWidth);
		//reference FreezeInfoLoaderImpl.getRowFreeze()
		sheet.getViewInfo().setNumOfRowFreeze(BookHelper.getRowFreeze(poiSheet));
		sheet.getViewInfo().setNumOfColumnFreeze(BookHelper.getColumnFreeze(poiSheet));
		sheet.getViewInfo().setDisplayGridline(poiSheet.isDisplayGridlines());
		sheet.getViewInfo().setColumnBreaks(poiSheet.getColumnBreaks());
		sheet.getViewInfo().setRowBreaks(poiSheet.getRowBreaks());
		
		NHeader header = sheet.getViewInfo().getHeader();
		header.setCenterText(poiSheet.getHeader().getCenter());
		header.setLeftText(poiSheet.getHeader().getLeft());
		header.setRightText(poiSheet.getHeader().getRight());
		
		NFooter footer = sheet.getViewInfo().getFooter();
		footer.setCenterText(poiSheet.getFooter().getCenter());
		footer.setLeftText(poiSheet.getFooter().getLeft());
		footer.setRightText(poiSheet.getFooter().getRight());
		
		sheet.getPrintSetup().setBottomMargin(UnitUtil.incheToPx(poiSheet.getMargin(Sheet.BottomMargin)));
		sheet.getPrintSetup().setTopMargin(UnitUtil.incheToPx(poiSheet.getMargin(Sheet.TopMargin)));
		sheet.getPrintSetup().setLeftMargin(UnitUtil.incheToPx(poiSheet.getMargin(Sheet.LeftMargin)));
		sheet.getPrintSetup().setRightMargin(UnitUtil.incheToPx(poiSheet.getMargin(Sheet.RightMargin)));
		
		sheet.getPrintSetup().setHeaderMargin(UnitUtil.incheToPx(poiSheet.getMargin(Sheet.HeaderMargin)));
		sheet.getPrintSetup().setFooterMargin(UnitUtil.incheToPx(poiSheet.getMargin(Sheet.FooterMargin)));
		sheet.getPrintSetup().setPaperSize(poiSheet.getPrintSetup().getPaperSize());
		sheet.getPrintSetup().setLandscape(poiSheet.getPrintSetup().getLandscape());
		sheet.getPrintSetup().setScale(poiSheet.getPrintSetup().getScale());
		
		sheet.setProtected(poiSheet.getProtect());
		
		return sheet;
	}

	protected void importMergedRegions(Sheet poiSheet, NSheet sheet) {
		//merged cells
		//reference RangeImpl.getMergeAreas()
		int nMerged = poiSheet.getNumMergedRegions();
		for(int i = nMerged - 1; i >= 0; --i) {
			final CellRangeAddress mergedRegion = poiSheet.getMergedRegion(i);
			sheet.addMergedRegion(new CellRegion(mergedRegion.getFirstRow(),mergedRegion.getFirstColumn(), mergedRegion.getLastRow(),mergedRegion.getLastColumn()));
		}
	}

	abstract protected void importDrawings(Sheet poiSheet, NSheet sheet);
	abstract protected void importValidation(Sheet poiSheet, NSheet sheet);
	
	protected NRow importRow(Row poiRow, NSheet sheet) {
		NRow row = sheet.getRow(poiRow.getRowNum());
		row.setHeight(UnitUtil.twipToPx(poiRow.getHeight()));
		row.setCustomHeight(poiRow.isCustomHeight());
		row.setHidden(poiRow.getZeroHeight());
		CellStyle rowStyle = poiRow.getRowStyle();
		if (rowStyle != null){
			row.setCellStyle(importCellStyle(rowStyle));
		}
		
		for(Cell poiCell : poiRow) {
			importCell(poiCell,poiRow.getRowNum(), sheet);
		}
		
		return row;
	}
	
	protected NCell importCell(Cell poiCell, int row, NSheet sheet){
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
			default:
				//TODO log: leave an unknown cell type as a blank cell.
				break;
		}
		cell.setCellStyle(importCellStyle(poiCell.getCellStyle()));
		
		return cell;
	}
	
	/**
	 * Convert CellStyle into NCellStyle
	 * @param poiCellStyle
	 */
	protected NCellStyle importCellStyle(CellStyle poiCellStyle) {
		NCellStyle cellStyle = null;
		if((cellStyle = importedStyle.get(poiCellStyle.getIndex())) ==null){
			cellStyle = book.createCellStyle(true);
			importedStyle.put(poiCellStyle.getIndex(), cellStyle);
			//reference XSSFDataFormat.getFormat()
			String dataFormat = null;
			if ((dataFormat = BuiltinFormats.getBuiltinFormat(poiCellStyle.getDataFormat()))==null){
				dataFormat = poiCellStyle.getDataFormatString();
			}
			cellStyle.setDataFormat(dataFormat);
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

	protected ErrorValue convertErrorCode(byte errorCellValue) {
		switch (errorCellValue){
			case ErrorConstants.ERROR_DIV_0:
				return new ErrorValue(ErrorValue.ERROR_DIV_0);
			case ErrorConstants.ERROR_NA:
				return new ErrorValue(ErrorValue.ERROR_NA);
			case ErrorConstants.ERROR_NAME:
				return new ErrorValue(ErrorValue.INVALID_NAME);
			case ErrorConstants.ERROR_NULL:
				return new ErrorValue(ErrorValue.ERROR_NULL);
			case ErrorConstants.ERROR_NUM:
				return new ErrorValue(ErrorValue.ERROR_NUM);
			case ErrorConstants.ERROR_REF:
				return new ErrorValue(ErrorValue.ERROR_REF);
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

	protected NViewAnchor toViewAnchor(Sheet poiSheet, ClientAnchor clientAnchor) {
		int width = getAnchorWidthInPx(clientAnchor, poiSheet);
		int height = getAnchorHeightInPx(clientAnchor, poiSheet);
		NViewAnchor viewAnchor = new NViewAnchor(clientAnchor.getRow1(), clientAnchor.getCol1(), width, height);
		viewAnchor.setXOffset(getXoffsetInPixel(clientAnchor, poiSheet));
		viewAnchor.setYOffset(getYoffsetInPixel(clientAnchor, poiSheet));
		return viewAnchor;
	}
	
	abstract protected int getXoffsetInPixel(ClientAnchor clientAnchor, Sheet poiSheet);
	abstract protected int getYoffsetInPixel(ClientAnchor clientAnchor, Sheet poiSheet);

	protected void importPicture(List<Picture> poiPictures, Sheet poiSheet, NSheet sheet) {
		for (Picture picture : poiPictures){
			Format format = Format.valueOfFileExtension(picture.getPictureData().suggestFileExtension());
			if (format !=null){
				sheet.addPicture(format, picture.getPictureData().getData(), toViewAnchor(poiSheet, picture.getClientAnchor()));
			}else{
				//TODO log we ignore a picture with unsupported format
			}
		}
	}
	
	/**
	 * A FilterColumn object only exists when we have set a criteria on that column. 
	 * For example, if we enable auto filter on 2 columns, but we only set criteria on 2nd column. 
	 * Thus, the size of filter column is 1. There is only one FilterColumn object and its column id is 1.  
	 * Only getFilterColumn(1) will return a FilterColumn, other get null.
	 * @param poiSheet
	 * @param sheet
	 */
	private void importAutoFilter(Sheet poiSheet, NSheet sheet) {
		AutoFilter poiAutoFilter = poiSheet.getAutoFilter();
		if (poiAutoFilter != null){
			CellRangeAddress filteringRange = poiAutoFilter.getRangeAddress();
			NAutoFilter autoFilter = sheet.createAutoFilter(new CellRegion(filteringRange.formatAsString()));
			int numberOfColumn = filteringRange.getLastColumn() - filteringRange.getFirstColumn() +1;
			for( int i = 0 ; i < numberOfColumn ; i ++){
				FilterColumn srcColumn = poiAutoFilter.getFilterColumn(i);
				if (srcColumn == null){
					continue;
				}
				NFilterColumn destColumn = autoFilter.getFilterColumn(i, true);
				destColumn.setProperties(toFilterOperator(srcColumn.getOperator()), srcColumn.getCriteria1(),
						srcColumn.getCriteria2(), srcColumn.isOn());
			}
		}
	}
	
	private FilterOp toFilterOperator(int operator){
		switch(operator){
			case AutoFilter.FILTEROP_AND:
				return FilterOp.AND;
			case AutoFilter.FILTEROP_BOTTOM10:
				return FilterOp.BOTTOM10;
			case AutoFilter.FILTEROP_BOTOOM10PERCENT:
				return FilterOp.BOTOOM10_PERCENT;
			case AutoFilter.FILTEROP_OR:
				return FilterOp.OR;
			case AutoFilter.FILTEROP_TOP10:
				return FilterOp.TOP10;
			case AutoFilter.FILTEROP_TOP10PERCENT:
				return FilterOp.TOP10_PERCENT;
			case AutoFilter.FILTEROP_VALUES: 
			default:
				return FilterOp.VALUES;
		}
	}
}
