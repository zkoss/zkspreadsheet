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
package org.zkoss.zss.range.impl.imexp;

import java.io.*;
import java.util.*;
import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.poi.ss.usermodel.*;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.util.logging.Log;
import org.zkoss.zss.model.*;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.SRichText.Segment;
import org.zkoss.zss.model.impl.sys.formula.FormulaEngineImpl;

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
	 * The map stores the exported {@link CellStyle} during exporting, so that
	 * we can reuse them for exporting other cells.
	 */
	protected Map<SCellStyle, CellStyle> styleTable = new HashMap<SCellStyle, CellStyle>();
	protected Map<SFont, Font> fontTable = new HashMap<SFont, Font>();
	protected Map<SColor, Color> colorTable = new HashMap<SColor, Color>();
	
	private static final Log _logger = Log.lookup(AbstractExcelExporter.class.getName());

	abstract protected void exportColumnArray(SSheet sheet, Sheet poiSheet, SColumnArray columnArr);

	abstract protected Workbook createPoiBook();

	abstract protected void exportChart(SSheet sheet, Sheet poiSheet);

	abstract protected void exportPicture(SSheet sheet, Sheet poiSheet);

	abstract protected void exportValidation(SSheet sheet, Sheet poiSheet);

	abstract protected void exportAutoFilter(SSheet sheet, Sheet poiSheet);

	/**
	 * Export the model according to reversed depended order: book, sheet,
	 * defined name, cells, chart, pictures, validation. Because named ranges
	 * (defined names) require sheet index, they should be imported after sheets
	 * created. Besides, cells, charts, and validations may have formulas
	 * referring to named ranges, they must be imported after named ranged.
	 * Pictures depend on cells.
	 */
	@Override
	public void export(SBook book, OutputStream fos) throws IOException {
		ReadWriteLock lock = book.getBookSeries().getLock();
		lock.readLock().lock();

		try {
			// clear cache for reuse
			styleTable.clear();
			fontTable.clear();
			colorTable.clear();
			
			workbook = createPoiBook();

			for (SSheet sheet : book.getSheets()) {
				exportSheet(sheet);
			}
			exportNamedRange(book);
			for (int n = 0; n < book.getSheets().size(); n++) {
				SSheet sheet = book.getSheet(n);
				Sheet poiSheet = workbook.getSheetAt(n);
				exportRowColumn(sheet, poiSheet);
				exportMergedRegions(sheet, poiSheet);
				exportChart(sheet, poiSheet);
				exportPicture(sheet, poiSheet);
				exportValidation(sheet, poiSheet);
				exportAutoFilter(sheet, poiSheet);
			}

			workbook.write(fos);
		} finally {
			lock.readLock().unlock();
		}
	}

	protected void exportNamedRange(SBook book) {
		for (SName name : book.getNames()) {
			Name poiName = workbook.createName();
			try{
				poiName.setNameName(name.getName());
				String sheetName = name.getRefersToSheetName();
				if(sheetName!=null){
					poiName.setSheetIndex(workbook.getSheetIndex(sheetName));
				}
				//zss-214, to tolerate the name refers to formula error (#REF!!$A$1:$I$18)
				if(!name.isFormulaParsingError()){
					poiName.setRefersToFormula(name.getRefersToFormula());
				}
			}catch (Exception e) {
				//ZSS-645 catch the exception happens when a book has a named range referring to an external book in XLS
				_logger.warning("Cannot export a name range: "+name.getName(),e);
				if (poiName.getNameName()!=null){
					workbook.removeName(poiName.getNameName());
				}
			}
			
		}
	}

	protected void exportSheet(SSheet sheet) {
		Sheet poiSheet = workbook.createSheet(sheet.getSheetName());

		// refer to BookHelper#setFreezePanel
		int freezeRow = sheet.getViewInfo().getNumOfRowFreeze();
		int freezeCol = sheet.getViewInfo().getNumOfColumnFreeze();
		poiSheet.createFreezePane(freezeCol <= 0 ? 0 : freezeCol, freezeRow <= 0 ? 0 : freezeRow);

		poiSheet.setDisplayGridlines(sheet.getViewInfo().isDisplayGridlines());

		if (sheet.isProtected()) {
			poiSheet.protectSheet(""); // without password
		}

		poiSheet.setDefaultRowHeight((short) UnitUtil.pxToTwip(sheet.getDefaultRowHeight()));
		poiSheet.setDefaultColumnWidth((int) UnitUtil.pxToDefaultColumnWidth(sheet.getDefaultColumnWidth(), AbstractExcelImporter.CHRACTER_WIDTH));

		// Header
		Header header = poiSheet.getHeader();
		header.setLeft(sheet.getViewInfo().getHeader().getLeftText());
		header.setCenter(sheet.getViewInfo().getHeader().getCenterText());
		header.setRight(sheet.getViewInfo().getHeader().getRightText());

		// Footer
		Footer footer = poiSheet.getFooter();
		footer.setLeft(sheet.getViewInfo().getFooter().getLeftText());
		footer.setCenter(sheet.getViewInfo().getFooter().getCenterText());
		footer.setRight(sheet.getViewInfo().getFooter().getRightText());

		// Margin
		poiSheet.setMargin(Sheet.LeftMargin, UnitUtil.pxToInche(sheet.getPrintSetup().getLeftMargin()));
		poiSheet.setMargin(Sheet.RightMargin, UnitUtil.pxToInche(sheet.getPrintSetup().getRightMargin()));
		poiSheet.setMargin(Sheet.TopMargin, UnitUtil.pxToInche(sheet.getPrintSetup().getTopMargin()));
		poiSheet.setMargin(Sheet.BottomMargin, UnitUtil.pxToInche(sheet.getPrintSetup().getBottomMargin()));

		// Print Setup Information
		poiSheet.getPrintSetup().setPaperSize(PoiEnumConversion.toPoiPaperSize(sheet.getPrintSetup().getPaperSize()));
		poiSheet.getPrintSetup().setLandscape(sheet.getPrintSetup().isLandscape());

	}

	protected void exportMergedRegions(SSheet sheet, Sheet poiSheet) {
		// consistent with importer, read from last merged region
		for (int i = sheet.getNumOfMergedRegion() - 1; i >= 0; i--) {
			CellRegion region = sheet.getMergedRegion(i);
			poiSheet.addMergedRegion(new CellRangeAddress(region.row, region.lastRow, region.column, region.lastColumn));
		}
	}

	protected void exportRowColumn(SSheet sheet, Sheet poiSheet) {
		// export rows
		Iterator<SRow> rowIterator = sheet.getRowIterator();
		while (rowIterator.hasNext()) {
			SRow row = rowIterator.next();
			exportRow(sheet, poiSheet, row);
		}

		// export columns
		Iterator<SColumnArray> columnArrayIterator = sheet.getColumnArrayIterator();
		while (columnArrayIterator.hasNext()) {
			SColumnArray columnArr = columnArrayIterator.next();
			exportColumnArray(sheet, poiSheet, columnArr);
		}
	}

	protected void exportRow(SSheet sheet, Sheet poiSheet, SRow row) {
		Row poiRow = poiSheet.createRow(row.getIndex());

		if (row.isHidden()) {
			// hidden, set height as 0
			poiRow.setZeroHeight(true);
		} else {
			// not hidden, calculate height
			if (row.getHeight() != sheet.getDefaultRowHeight()) {
				poiRow.setCustomHeight(true);
				poiRow.setHeight((short) UnitUtil.pxToTwip(row.getHeight()));
			}
		}

		SCellStyle rowStyle = row.getCellStyle();
		CellStyle poiRowStyle = toPOICellStyle(rowStyle);
		poiRow.setRowStyle(poiRowStyle);

		// Export Cell
		Iterator<SCell> cellIterator = sheet.getCellIterator(row.getIndex());
		while (cellIterator.hasNext()) {
			SCell cell = cellIterator.next();
			exportCell(poiRow, cell);
		}
	}

	protected void exportCell(Row poiRow, SCell cell) {
		Cell poiCell = poiRow.createCell(cell.getColumnIndex());

		SCellStyle cellStyle = cell.getCellStyle();
		poiCell.setCellStyle(toPOICellStyle(cellStyle));

		switch (cell.getType()) {
		case BLANK:
			poiCell.setCellType(Cell.CELL_TYPE_BLANK);
			break;
		case ERROR:
			//ignore the value of this cell, excel doesn't allow it invalid formula (pasring error).
			if(cell.getErrorValue().getCode() != ErrorValue.INVALID_FORMULA){
				poiCell.setCellType(Cell.CELL_TYPE_ERROR);
				poiCell.setCellErrorValue(cell.getErrorValue().getCode());
			}
			break;
		case BOOLEAN:
			poiCell.setCellType(Cell.CELL_TYPE_BOOLEAN);
			poiCell.setCellValue(cell.getBooleanValue());
			break;
		case FORMULA:
			if(cell.getFormulaResultType()==CellType.ERROR && cell.getErrorValue().getCode() != ErrorValue.INVALID_FORMULA){
				//ignore the value of this cell, excel doesn't allow it invalid formula (pasring error).
			}else{
				poiCell.setCellType(Cell.CELL_TYPE_FORMULA);
				poiCell.setCellFormula(cell.getFormulaValue());
			}
			break;
		case NUMBER:
			poiCell.setCellType(Cell.CELL_TYPE_NUMERIC);
			poiCell.setCellValue((Double) cell.getNumberValue());
			break;
		case STRING:
			poiCell.setCellType(Cell.CELL_TYPE_STRING);
			if(cell.isRichTextValue()) {
				poiCell.setCellValue(toPOIRichText(cell.getRichTextValue()));
			} else {
				poiCell.setCellValue(cell.getStringValue());
			}
			break;
		default:
		}
		
		SHyperlink hyperlink = cell.getHyperlink();
		if (hyperlink != null) {
			CreationHelper helper = workbook.getCreationHelper();
			try{
				Hyperlink poiHyperlink = helper.createHyperlink(PoiEnumConversion.toPoiHyperlinkType(hyperlink.getType()));
				poiHyperlink.setAddress(hyperlink.getAddress());
				poiHyperlink.setLabel(hyperlink.getLabel());
				poiCell.setHyperlink(poiHyperlink);
			}catch (Exception e) {
				//ZSS-644 catch the exception happens when a hyperlink has an invalid URI in XLSX
				_logger.warning("Cannot export a hyperlink: "+hyperlink.getAddress(),e);
			}
			
		}
		
		SComment comment = cell.getComment();
		if (comment != null) {
			// Refer to the POI Official Tutorial
			// http://poi.apache.org/spreadsheet/quick-guide.html#CellComments
			CreationHelper helper = workbook.getCreationHelper();
			Drawing drawing = poiCell.getSheet().createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();
			anchor.setCol1(poiCell.getColumnIndex());
			anchor.setCol2(poiCell.getColumnIndex() + 1);
			anchor.setRow1(poiRow.getRowNum());
			anchor.setRow2(poiRow.getRowNum() + 3);
			Comment poiComment = drawing.createCellComment(anchor);
			SRichText richText = comment.getRichText();
			if (richText != null) {
				poiComment.setString(toPOIRichText(richText));
			} else {
				poiComment.setString(helper.createRichTextString(comment.getText()));
			}
			poiComment.setAuthor(comment.getAuthor());
			poiComment.setVisible(comment.isVisible());
			poiCell.setCellComment(poiComment);
		}	
	}
	
	protected RichTextString toPOIRichText(SRichText richText) {

		CreationHelper helper = workbook.getCreationHelper();

		RichTextString poiRichTextString = helper.createRichTextString(richText.getText());

		int start = 0;
		int end = 0;
		for (Segment sg : richText.getSegments()) {
			SFont font = sg.getFont();
			int len = sg.getText().length();
			end += len;
			poiRichTextString.applyFont(start, end, toPOIFont(font));
			start += len;
		}

		return poiRichTextString;
	}

	protected CellStyle toPOICellStyle(SCellStyle cellStyle) {
		// instead of creating a new style, use old one if exist
		CellStyle poiCellStyle = styleTable.get(cellStyle);
		if (poiCellStyle != null) {
			return poiCellStyle;
		}

		poiCellStyle = workbook.createCellStyle();

		/* Bottom Border */
		poiCellStyle.setBorderBottom(PoiEnumConversion.toPoiBorderType(cellStyle.getBorderBottom()));

		BookHelper.setBottomBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderBottomColor()));

		/* Left Border */
		poiCellStyle.setBorderLeft(PoiEnumConversion.toPoiBorderType(cellStyle.getBorderLeft()));
		BookHelper.setLeftBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderLeftColor()));

		/* Right Border */
		poiCellStyle.setBorderRight(PoiEnumConversion.toPoiBorderType(cellStyle.getBorderRight()));
		BookHelper.setRightBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderRightColor()));

		/* Top Border */
		poiCellStyle.setBorderTop(PoiEnumConversion.toPoiBorderType(cellStyle.getBorderTop()));
		BookHelper.setTopBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderTopColor()));

		/* Fill Foreground Color */
		BookHelper.setFillForegroundColor(poiCellStyle, toPOIColor(cellStyle.getFillColor()));

		poiCellStyle.setFillPattern(PoiEnumConversion.toPoiFillPattern(cellStyle.getFillPattern()));
		poiCellStyle.setAlignment(PoiEnumConversion.toPoiHorizontalAlignment(cellStyle.getAlignment()));
		poiCellStyle.setVerticalAlignment(PoiEnumConversion.toPoiVerticalAlignment(cellStyle.getVerticalAlignment()));
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

	protected Color toPOIColor(SColor color) {
		Color poiColor = colorTable.get(color);

		if (poiColor != null) {
			return poiColor;
		}
		poiColor = BookHelper.HTMLToColor(workbook, color.getHtmlColor());
		colorTable.put(color, poiColor);
		return poiColor;
	}

	/**
	 * Convert ZSS Font into POI Font. Cache font in the fontTable. If font
	 * exist, don't create a new one.
	 * 
	 * @param font
	 * @return
	 */
	protected Font toPOIFont(SFont font) {

		Font poiFont = fontTable.get(font);
		if (poiFont != null) {
			return poiFont;
		}

		poiFont = workbook.createFont();
		poiFont.setBoldweight(PoiEnumConversion.toPoiBoldweight(font.getBoldweight()));
		poiFont.setStrikeout(font.isStrikeout());
		poiFont.setItalic(font.isItalic());
		BookHelper.setFontColor(workbook, poiFont, toPOIColor(font.getColor()));
		poiFont.setFontHeightInPoints((short) font.getHeightPoints());
		poiFont.setFontName(font.getName());
		poiFont.setTypeOffset(PoiEnumConversion.toPoiTypeOffset(font.getTypeOffset()));
		poiFont.setUnderline(PoiEnumConversion.toPoiUnderline(font.getUnderline()));

		// put into table
		fontTable.put(font, poiFont);

		return poiFont;
	}

}
