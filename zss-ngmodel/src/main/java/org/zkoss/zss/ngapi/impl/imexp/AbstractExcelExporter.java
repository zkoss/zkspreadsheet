/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/18 , Created by Hawk
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi.impl.imexp;

import java.io.IOException;
import java.io.OutputStream;
import java.util.*;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.ngmodel.*;
import org.zkoss.zss.ngmodel.NCellStyle.Alignment;
import org.zkoss.zss.ngmodel.NCellStyle.BorderType;
import org.zkoss.zss.ngmodel.NCellStyle.FillPattern;
import org.zkoss.zss.ngmodel.NCellStyle.VerticalAlignment;
import org.zkoss.zss.ngmodel.NFont.Boldweight;
import org.zkoss.zss.ngmodel.NFont.TypeOffset;
import org.zkoss.zss.ngmodel.NFont.Underline;

/**
 * Common exporting behavior for both XLSX and XLS.
 * 
 * @author kuro, Hawk
 * @since 3.5.0
 */
abstract public class AbstractExcelExporter extends AbstractExporter {
	
	/**
	 * Exporting destination, POI book model 
	 */
	protected Workbook workbook;
	/**
	 * The map stores the exported {@link CellStyle} during exporting, so that we can reuse them for exporting other cells.
	 */
	protected Map<NCellStyle, CellStyle> styleTable = new HashMap<NCellStyle, CellStyle>();
	protected Map<NFont, Font> fontTable = new HashMap<NFont, Font>();
	protected Map<NColor, Color> colorTable = new HashMap<NColor, Color>();
	
	abstract protected void exportColumnArray(NSheet sheet, Sheet poiSheet, NColumnArray columnArr);
	abstract protected Workbook createPoiBook();
	abstract protected void exportChart(NSheet sheet, Sheet poiSheet);
	abstract protected void exportPicture(NSheet sheet, Sheet poiSheet);
	
	@Override
	public void export(NBook book, OutputStream fos) throws IOException {
		ReadWriteLock lock = book.getBookSeries().getLock();
		lock.readLock().lock();
		
		try {
			
			workbook = createPoiBook();
			
			for(NSheet sheet : book.getSheets()) {
				exportSheet(sheet);
			}
			
			exportNamedRange(book);
			
			workbook.write(fos);
		} finally {
			lock.readLock().unlock();
		}
	}
	
	
	protected void exportNamedRange(NBook book) {
		
		for(NName name : book.getNames()) {
			Name poiName = workbook.createName();
			poiName.setNameName(name.getName());
			poiName.setSheetIndex(workbook.getSheetIndex(name.getRefersToSheetName()));
			poiName.setRefersToFormula(name.getRefersToFormula());
		}
	}
	
	protected void exportSheet(NSheet sheet) {
		Sheet poiSheet = workbook.createSheet(sheet.getSheetName());

		// consistent with importer, read from last merged region
		for(int i = sheet.getNumOfMergedRegion() - 1; i >= 0; i--) {
			CellRegion region = sheet.getMergedRegion(i);
			poiSheet.addMergedRegion(new CellRangeAddress(region.row, region.lastRow, region.column, region.lastColumn));
		}
		
		// refer to BookHelper#setFreezePanel
		int freezeRow = sheet.getViewInfo().getNumOfRowFreeze();
		int freezeCol = sheet.getViewInfo().getNumOfColumnFreeze();
		poiSheet.createFreezePane(freezeCol <= 0 ? 0 : freezeCol, freezeRow <= 0 ? 0 : freezeRow);

		poiSheet.setDisplayGridlines(sheet.getViewInfo().isDisplayGridline());
		
		if(sheet.isProtected()) {
			poiSheet.protectSheet(""); // without password
		}
		
		poiSheet.setDefaultRowHeight((short)XUtils.pxToTwip(sheet.getDefaultRowHeight()));
		poiSheet.setDefaultColumnWidth((int)XUtils.pxToDefaultColumnWidth(sheet.getDefaultColumnWidth(), AbstractExcelImporter.CHRACTER_WIDTH));
		//poiSheet.setDefaultColumnWidth((int)XUtils.pxToCTChar(sheet.getDefaultColumnWidth(), AbstractExcelImporter.CHRACTER_WIDTH));
		
		exportChart(sheet, poiSheet);
		exportPicture(sheet, poiSheet);
		
		//export rows
		Iterator<NRow> rowIterator = sheet.getRowIterator();
		while(rowIterator.hasNext()) {
			NRow row = rowIterator.next();
			exportRow(sheet, poiSheet, row);
		} 
		
		//export columns
		Iterator<NColumnArray> columnArrayIterator = sheet.getColumnArrayIterator();
		while(columnArrayIterator.hasNext()) {
			NColumnArray columnArr = columnArrayIterator.next();
			exportColumnArray(sheet, poiSheet, columnArr);
		}

	}

	protected void exportRow(NSheet sheet, Sheet poiSheet, NRow row) {
		Row poiRow = poiSheet.createRow(row.getIndex());
		
		if(row.isHidden()) {
			// hidden, set height as 0
			poiRow.setZeroHeight(true);
		} else {
			// not hidden, calculate height
			if(row.getHeight() != sheet.getDefaultRowHeight()) {
				poiRow.setCustomHeight(true);
				poiRow.setHeight((short)XUtils.pxToTwip(row.getHeight()));
			}
		}
		
		NCellStyle rowStyle = row.getCellStyle();
		CellStyle poiRowStyle = toPOICellStyle(rowStyle);
		poiRow.setRowStyle(poiRowStyle);
		
		// Export Cell
		Iterator<NCell> cellIterator = sheet.getCellIterator(row.getIndex());
		while(cellIterator.hasNext()) {
			NCell cell = cellIterator.next();
			exportCell(poiRow, cell);
		}
	}

	protected void exportCell(Row poiRow, NCell cell) {
		Cell poiCell = poiRow.createCell(cell.getColumnIndex());
		
		NCellStyle cellStyle = cell.getCellStyle();

		poiCell.setCellStyle(toPOICellStyle(cellStyle));
		
		switch(cell.getType()) {
			case BLANK:
				poiCell.setCellType(Cell.CELL_TYPE_BLANK);
				break;
			case ERROR:
				poiCell.setCellType(Cell.CELL_TYPE_ERROR);
				poiCell.setCellErrorValue(cell.getErrorValue().getCode());
				break;
			case BOOLEAN:
				poiCell.setCellType(Cell.CELL_TYPE_BOOLEAN);
				poiCell.setCellValue(cell.getBooleanValue());
				break;
			case FORMULA:
				poiCell.setCellType(Cell.CELL_TYPE_FORMULA);
				poiCell.setCellFormula(cell.getFormulaValue());
				break;
			case NUMBER:
				poiCell.setCellType(Cell.CELL_TYPE_NUMERIC);
				poiCell.setCellValue((Double)cell.getNumberValue());
				break;
			case STRING:
				poiCell.setCellType(Cell.CELL_TYPE_STRING);
				poiCell.setCellValue(cell.getStringValue());
				break;
			default:
		}
		
		// Hyperlink
//		NHyperlink hyperlink = null;
//		if((hyperlink = cell.getHyperlink()) != null) {
//			XSSFHyperlink xssfLink = xssfSheet.getHyperlink(cell.getRowIndex(), cell.getColumnIndex());
//			xssfLink.setAddress(hyperlink.getAddress());
//			xssfLink.setLabel(hyperlink.getLabel());
//			// TODO, some API isn't available.
//		}

	}
	
	protected CellStyle toPOICellStyle(NCellStyle cellStyle) {
		// instead of creating a new style, use old one if exist
		CellStyle poiCellStyle = styleTable.get(cellStyle);
		if(poiCellStyle != null) {
			return poiCellStyle;
		}
		
		poiCellStyle = workbook.createCellStyle();
		
		/* Bottom Border */
		poiCellStyle.setBorderBottom(toPOIBorderType(cellStyle.getBorderBottom()));
		
		BookHelper.setBottomBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderBottomColor()));
		
		/* Left Border */
		poiCellStyle.setBorderLeft(toPOIBorderType(cellStyle.getBorderLeft()));
		BookHelper.setLeftBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderLeftColor()));
		
		/* Right Border */
		poiCellStyle.setBorderRight(toPOIBorderType(cellStyle.getBorderRight()));
		BookHelper.setRightBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderRightColor()));
		
		/* Top Border*/
		poiCellStyle.setBorderTop(toPOIBorderType(cellStyle.getBorderTop()));
		BookHelper.setTopBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderTopColor()));
		
		/* Fill Foreground Color */
		BookHelper.setFillForegroundColor(poiCellStyle, toPOIColor(cellStyle.getFillColor()));
		
		poiCellStyle.setFillPattern(toPOIFillPattern(cellStyle.getFillPattern()));
		poiCellStyle.setAlignment(toPOIAlignment(cellStyle.getAlignment()));
		poiCellStyle.setVerticalAlignment(toPOIVerticalAlignment(cellStyle.getVerticalAlignment()));
		poiCellStyle.setWrapText(cellStyle.isWrapText());
		poiCellStyle.setLocked(cellStyle.isLocked());
		poiCellStyle.setHidden(cellStyle.isHidden());
		
		// refer from BookHelper#setDataFormat
		DataFormat df = workbook.createDataFormat();
		short fmt = df.getFormat(cellStyle.getDataFormat());
		poiCellStyle.setDataFormat(fmt);
		
		// font
		poiCellStyle.setFont(toPOIFont(cellStyle.getFont()));
		
		// put into table
		styleTable.put(cellStyle, poiCellStyle);
		
		return poiCellStyle;
		
	}
	
	protected Color toPOIColor(NColor color) {
		Color poiColor = colorTable.get(color);
		
		if(poiColor != null) {
			return poiColor;
		}
		poiColor = BookHelper.HTMLToColor(workbook, color.getHtmlColor());
		colorTable.put(color, poiColor);
		return poiColor;
	}
	
	protected Font toPOIFont(NFont font) {

		Font poiFont = fontTable.get(font);
		if(poiFont != null) {
			return poiFont;
		}
		
		poiFont = workbook.createFont();
		poiFont.setBoldweight(toPOIBoldweight(font.getBoldweight()));
		poiFont.setStrikeout(font.isStrikeout());
		poiFont.setItalic(font.isItalic());
		BookHelper.setFontColor(workbook, poiFont, toPOIColor(font.getColor()));
		poiFont.setFontHeightInPoints((short)font.getHeightPoints());
		poiFont.setFontName(font.getName());
		poiFont.setTypeOffset(toPOITypeOffset(font.getTypeOffset()));
		poiFont.setUnderline(toPOIUnderline(font.getUnderline()));
		
		// put into table
		fontTable.put(font, poiFont);
		
		return poiFont;
	}
	
	protected short toPOIVerticalAlignment(VerticalAlignment vAlignment) {
		switch(vAlignment) {
			case BOTTOM:
				return CellStyle.VERTICAL_BOTTOM;
			case CENTER:
				return CellStyle.VERTICAL_CENTER;
			case JUSTIFY:
				return CellStyle.VERTICAL_JUSTIFY;
			case TOP:
			default:
				return CellStyle.VERTICAL_TOP;
		}
	}
	
	protected short toPOIFillPattern(FillPattern fillPattern) {
		switch(fillPattern) {
		case ALT_BARS:
			return CellStyle.ALT_BARS;
		case BIG_SPOTS:
			return CellStyle.BIG_SPOTS;
		case BRICKS:
			return CellStyle.BRICKS;
		case DIAMONDS:
			return CellStyle.DIAMONDS;
		case FINE_DOTS:
			return CellStyle.FINE_DOTS;
		case LEAST_DOTS:
			return CellStyle.LEAST_DOTS;
		case LESS_DOTS:
			return CellStyle.LESS_DOTS;
		case SOLID_FOREGROUND:
			return CellStyle.SOLID_FOREGROUND;
		case SPARSE_DOTS:
			return CellStyle.SPARSE_DOTS;
		case SQUARES:
			return CellStyle.SQUARES;
		case THICK_BACKWARD_DIAG:
			return CellStyle.THICK_BACKWARD_DIAG;
		case THICK_FORWARD_DIAG:
			return CellStyle.THICK_FORWARD_DIAG;
		case THICK_HORZ_BANDS:
			return CellStyle.THICK_HORZ_BANDS;
		case THICK_VERT_BANDS:
			return CellStyle.THICK_VERT_BANDS;
		case THIN_BACKWARD_DIAG:
			return CellStyle.THIN_BACKWARD_DIAG;
		case THIN_FORWARD_DIAG:
			return CellStyle.THIN_FORWARD_DIAG;
		case THIN_HORZ_BANDS:
			return CellStyle.THIN_HORZ_BANDS;
		case THIN_VERT_BANDS:
			return CellStyle.THIN_VERT_BANDS;
		case NO_FILL:
		default:
			return CellStyle.NO_FILL;
		}
	}
	
	protected short toPOIBorderType(BorderType borderType) {
		switch(borderType) {
		case DASH_DOT:
			return CellStyle.BORDER_DASH_DOT;
		case DASHED:
			return CellStyle.BORDER_DASHED;
		case DOTTED:
			return CellStyle.BORDER_DOTTED;
		case DOUBLE:
			return CellStyle.BORDER_DOUBLE;
		case HAIR:
			return CellStyle.BORDER_HAIR;
		case MEDIUM:
			return CellStyle.BORDER_MEDIUM;
		case MEDIUM_DASH_DOT:
			return CellStyle.BORDER_DASH_DOT;
		case MEDIUM_DASH_DOT_DOT:
			return CellStyle.BORDER_DASH_DOT_DOT;
		case MEDIUM_DASHED:
			return CellStyle.BORDER_MEDIUM_DASHED;
		case SLANTED_DASH_DOT:
			return CellStyle.BORDER_SLANTED_DASH_DOT;
		case THICK:
			return CellStyle.BORDER_THICK;
		case THIN:
			return CellStyle.BORDER_THIN;
		case DASH_DOT_DOT:
			return CellStyle.BORDER_DASH_DOT_DOT;
		case NONE:
		default:
			return CellStyle.BORDER_NONE;
		}
	}
	
	protected short toPOIAlignment(Alignment alignment) {
		switch(alignment) {
		case CENTER:
			return CellStyle.ALIGN_CENTER;
		case FILL:
			return CellStyle.ALIGN_FILL;
		case JUSTIFY:
			return CellStyle.ALIGN_JUSTIFY;
		case RIGHT:
			return CellStyle.ALIGN_RIGHT;
		case LEFT:
			return CellStyle.ALIGN_LEFT;
		case CENTER_SELECTION:
			return CellStyle.ALIGN_CENTER_SELECTION;
		case GENERAL:
			default:
			return CellStyle.ALIGN_GENERAL;
		}
	}
	
	protected short toPOIBoldweight(Boldweight bold) {
		switch(bold) {
			case BOLD:
				return Font.BOLDWEIGHT_BOLD;
			case NORMAL:
			default:
				return Font.BOLDWEIGHT_NORMAL;
		}
	}
	
	protected short toPOITypeOffset(TypeOffset typeOffset) {
		switch(typeOffset) {
			case SUPER:
				return Font.SS_SUPER;
			case SUB:
				return Font.SS_SUB;
			case NONE:
			default:
				return Font.SS_NONE;
		}
	}
	
	protected byte toPOIUnderline(Underline underline) {
		switch(underline) {
			case SINGLE:
				return Font.U_SINGLE;
			case DOUBLE:
				return Font.U_DOUBLE;
			case DOUBLE_ACCOUNTING:
				return Font.U_DOUBLE_ACCOUNTING;
			case SINGLE_ACCOUNTING:
				return Font.U_SINGLE_ACCOUNTING;
			case NONE:
			default:
				return Font.U_NONE;
		}
	}
}
