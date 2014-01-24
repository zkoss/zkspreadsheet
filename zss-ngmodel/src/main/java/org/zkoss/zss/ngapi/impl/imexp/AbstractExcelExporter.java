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
import org.zkoss.zss.ngmodel.NRichText.Segment;

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
	protected Map<NCellStyle, CellStyle> styleTable = new HashMap<NCellStyle, CellStyle>();
	protected Map<NFont, Font> fontTable = new HashMap<NFont, Font>();
	protected Map<NColor, Color> colorTable = new HashMap<NColor, Color>();

	abstract protected void exportColumnArray(NSheet sheet, Sheet poiSheet, NColumnArray columnArr);

	abstract protected Workbook createPoiBook();

	abstract protected void exportChart(NSheet sheet, Sheet poiSheet);

	abstract protected void exportPicture(NSheet sheet, Sheet poiSheet);

	abstract protected void exportValidation(NSheet sheet, Sheet poiSheet);

	abstract protected void exportAutoFilter(NSheet sheet, Sheet poiSheet);

	/**
	 * Export the model according to reversed depended order: book, sheet,
	 * defined name, cells, chart, pictures, validation. Because named ranges
	 * (defined names) require sheet index, they should be imported after sheets
	 * created. Besides, cells, charts, and validations may have formulas
	 * referring to named ranges, they must be imported after named ranged.
	 * Pictures depend on cells.
	 */
	@Override
	public void export(NBook book, OutputStream fos) throws IOException {
		ReadWriteLock lock = book.getBookSeries().getLock();
		lock.readLock().lock();

		try {
			workbook = createPoiBook();

			for (NSheet sheet : book.getSheets()) {
				exportSheet(sheet);
			}
			exportNamedRange(book);
			for (int n = 0; n < book.getSheets().size(); n++) {
				NSheet sheet = book.getSheet(n);
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

	protected void exportNamedRange(NBook book) {

		for (NName name : book.getNames()) {
			Name poiName = workbook.createName();
			poiName.setNameName(name.getName());
			poiName.setSheetIndex(workbook.getSheetIndex(name.getRefersToSheetName()));
			poiName.setRefersToFormula(name.getRefersToFormula());
		}
	}

	protected void exportSheet(NSheet sheet) {
		Sheet poiSheet = workbook.createSheet(sheet.getSheetName());

		// refer to BookHelper#setFreezePanel
		int freezeRow = sheet.getViewInfo().getNumOfRowFreeze();
		int freezeCol = sheet.getViewInfo().getNumOfColumnFreeze();
		poiSheet.createFreezePane(freezeCol <= 0 ? 0 : freezeCol, freezeRow <= 0 ? 0 : freezeRow);

		poiSheet.setDisplayGridlines(sheet.getViewInfo().isDisplayGridline());

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
		poiSheet.getPrintSetup().setPaperSize(ExporterEnumUtil.toPoiPaperSize(sheet.getPrintSetup().getPaperSize()));
		poiSheet.getPrintSetup().setLandscape(sheet.getPrintSetup().isLandscape());

	}

	protected void exportMergedRegions(NSheet sheet, Sheet poiSheet) {
		// consistent with importer, read from last merged region
		for (int i = sheet.getNumOfMergedRegion() - 1; i >= 0; i--) {
			CellRegion region = sheet.getMergedRegion(i);
			poiSheet.addMergedRegion(new CellRangeAddress(region.row, region.lastRow, region.column, region.lastColumn));
		}
	}

	protected void exportRowColumn(NSheet sheet, Sheet poiSheet) {
		// export rows
		Iterator<NRow> rowIterator = sheet.getRowIterator();
		while (rowIterator.hasNext()) {
			NRow row = rowIterator.next();
			exportRow(sheet, poiSheet, row);
		}

		// export columns
		Iterator<NColumnArray> columnArrayIterator = sheet.getColumnArrayIterator();
		while (columnArrayIterator.hasNext()) {
			NColumnArray columnArr = columnArrayIterator.next();
			exportColumnArray(sheet, poiSheet, columnArr);
		}
	}

	protected void exportRow(NSheet sheet, Sheet poiSheet, NRow row) {
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

		NCellStyle rowStyle = row.getCellStyle();
		CellStyle poiRowStyle = toPOICellStyle(rowStyle);
		poiRow.setRowStyle(poiRowStyle);

		// Export Cell
		Iterator<NCell> cellIterator = sheet.getCellIterator(row.getIndex());
		while (cellIterator.hasNext()) {
			NCell cell = cellIterator.next();
			exportCell(poiRow, cell);
		}
	}

	protected void exportCell(Row poiRow, NCell cell) {
		Cell poiCell = poiRow.createCell(cell.getColumnIndex());

		NCellStyle cellStyle = cell.getCellStyle();
		poiCell.setCellStyle(toPOICellStyle(cellStyle));

		switch (cell.getType()) {
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
			poiCell.setCellValue((Double) cell.getNumberValue());
			break;
		case STRING:
			poiCell.setCellType(Cell.CELL_TYPE_STRING);
			if(cell.isRichTextValue()) {
				NRichText richText = cell.getRichTextValue();
				CreationHelper helper = workbook.getCreationHelper();
				RichTextString poiRichTextString = helper.createRichTextString(richText.getText());
				int start = 0;
				int end = 0;
				for(Segment sg : richText.getSegments()) {
					NFont font = sg.getFont();
					int len = sg.getText().length();
					end += len;
					poiRichTextString.applyFont(start, end, toPOIFont(font));
					start += len;
				}
				poiCell.setCellValue(poiRichTextString);
			} else {
				poiCell.setCellValue(cell.getStringValue());
			}
			break;
		default:
		}
		
		NHyperlink hyperlink = cell.getHyperlink();
		if (hyperlink != null) {
			CreationHelper helper = workbook.getCreationHelper();
			Hyperlink poiHyperlink = helper.createHyperlink(ExporterEnumUtil.toPoiHyperlinkType(hyperlink.getType()));
			poiHyperlink.setAddress(hyperlink.getAddress());
			poiHyperlink.setLabel(hyperlink.getLabel());
			poiCell.setHyperlink(poiHyperlink);
		}
	}

	protected CellStyle toPOICellStyle(NCellStyle cellStyle) {
		// instead of creating a new style, use old one if exist
		CellStyle poiCellStyle = styleTable.get(cellStyle);
		if (poiCellStyle != null) {
			return poiCellStyle;
		}

		poiCellStyle = workbook.createCellStyle();

		/* Bottom Border */
		poiCellStyle.setBorderBottom(ExporterEnumUtil.toPoiBorderType(cellStyle.getBorderBottom()));

		BookHelper.setBottomBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderBottomColor()));

		/* Left Border */
		poiCellStyle.setBorderLeft(ExporterEnumUtil.toPoiBorderType(cellStyle.getBorderLeft()));
		BookHelper.setLeftBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderLeftColor()));

		/* Right Border */
		poiCellStyle.setBorderRight(ExporterEnumUtil.toPoiBorderType(cellStyle.getBorderRight()));
		BookHelper.setRightBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderRightColor()));

		/* Top Border */
		poiCellStyle.setBorderTop(ExporterEnumUtil.toPoiBorderType(cellStyle.getBorderTop()));
		BookHelper.setTopBorderColor(poiCellStyle, toPOIColor(cellStyle.getBorderTopColor()));

		/* Fill Foreground Color */
		BookHelper.setFillForegroundColor(poiCellStyle, toPOIColor(cellStyle.getFillColor()));

		poiCellStyle.setFillPattern(ExporterEnumUtil.toPoiFillPattern(cellStyle.getFillPattern()));
		poiCellStyle.setAlignment(ExporterEnumUtil.toPoiAlignment(cellStyle.getAlignment()));
		poiCellStyle.setVerticalAlignment(ExporterEnumUtil.toPoiVerticalAlignment(cellStyle.getVerticalAlignment()));
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
	protected Font toPOIFont(NFont font) {

		Font poiFont = fontTable.get(font);
		if (poiFont != null) {
			return poiFont;
		}

		poiFont = workbook.createFont();
		poiFont.setBoldweight(ExporterEnumUtil.toPoiBoldweight(font.getBoldweight()));
		poiFont.setStrikeout(font.isStrikeout());
		poiFont.setItalic(font.isItalic());
		BookHelper.setFontColor(workbook, poiFont, toPOIColor(font.getColor()));
		poiFont.setFontHeightInPoints((short) font.getHeightPoints());
		poiFont.setFontName(font.getName());
		poiFont.setTypeOffset(ExporterEnumUtil.toPoiTypeOffset(font.getTypeOffset()));
		poiFont.setUnderline(ExporterEnumUtil.toPoiUnderline(font.getUnderline()));

		// put into table
		fontTable.put(font, poiFont);

		return poiFont;
	}

}
