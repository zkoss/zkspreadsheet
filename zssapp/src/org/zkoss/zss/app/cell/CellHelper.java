package org.zkoss.zss.app.cell;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Font;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.impl.Utils;

public final class CellHelper {
	private CellHelper(){};
	
	/**
	 * Returns cell font family
	 * @param cell
	 * @return
	 */
	public static Font getFont(Cell cell) {
		if (cell == null)
			return null;
		
		Book book = (Book)cell.getSheet().getWorkbook();
		CellStyle cellStyle = cell.getCellStyle();
		return book.getFontAt(cellStyle.getFontIndex());
	}
	
	public static boolean isBold(Font font) {
		return font.getBoldweight() == Font.BOLDWEIGHT_BOLD;
	}
	
	public static boolean isAlignLeft(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		return cellStyle.getAlignment() == CellStyle.ALIGN_LEFT;
	}
	
	public static boolean isAlignCenter(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		return cellStyle.getAlignment() == CellStyle.ALIGN_CENTER;
	}

	public static boolean isAlignRight(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		return cellStyle.getAlignment() == CellStyle.ALIGN_RIGHT;
	}
	
	public static String getFontHTMLColor(Cell cell, Font font) {
		Book book = (Book)cell.getSheet().getWorkbook();
		String color = BookHelper.getFontHTMLColor(book, font);
		if (color == null || BookHelper.AUTO_COLOR.equals(color))
			return "#000000";
		return  color;
	}
	
	public static String getBackgroundHTMLColor(Cell cell) {
		CellStyle cellStyle = cell.getCellStyle();
		Book book = (Book)cell.getSheet().getWorkbook();
		String color = BookHelper.colorToHTML(book, cellStyle.getFillForegroundColorColor());
		if (color == null || BookHelper.AUTO_COLOR.equals(color))
			return "#FFFFFF";
		return color;
	}

	
	/**
	 * Delete current cell and shift cells up
	 * @param cell
	 */
	public static void shiftCellUp(Sheet sheet, int rowIndex, int colIndex) {
		final Range rng = Ranges.range(sheet, 
				rowIndex, 
				colIndex);
		rng.delete(Range.SHIFT_UP);
	}
	
	/**
	 * Insert a new cell and shift cells right
	 * @param cell
	 */
	public static void shiftCellRight(Sheet sheet, int rowIndex, int colIndex) {
		final Range rng = Ranges.range(sheet, 
				rowIndex, 
				colIndex);
		rng.insert(Range.SHIFT_RIGHT, Range.FORMAT_RIGHTBELOW);
	}
	
	/**
	 * Insert a new cell and shift original cells down
	 */
	public static void shiftCellDown(Sheet sheet, int rowIndex, int colIndex) {
		final Range rng = Ranges.range(sheet, 
				rowIndex, 
				colIndex);
		rng.insert(Range.SHIFT_DOWN, Range.FORMAT_LEFTABOVE);
	}
	
	/**
	 * Delete current cell and shift cells up beside it.
	 * @param cell
	 */
	public static void shiftCellLeft(Sheet sheet, int rowIndex, int colIndex) {
		final Range rng = Ranges.range(sheet, 
				rowIndex, 
				colIndex);
		rng.delete(Range.SHIFT_LEFT);
	}
	
	public static void shiftEntireRowDown(Sheet sheet, int rowIndex, int colIndex) {
		Row row = sheet.getRow(rowIndex);
		int lCol = row.getFirstCellNum();
		int rCol  = row.getLastCellNum();
		for(int colIdx = lCol; colIdx < rCol; colIdx++) {
			final Range rng = Ranges.range(sheet, rowIndex,	colIdx);
			rng.insert(Range.SHIFT_DOWN, Range.FORMAT_LEFTABOVE);
		}
	}
	
	public static void shiftEntireRowUp(Sheet sheet, int rowIndex, int colIndex) {
		
		Row row = sheet.getRow(rowIndex);
		int lCol = row.getFirstCellNum();
		int rCol  = row.getLastCellNum();
		for(int colIdx = lCol; colIdx < rCol; colIdx++) {
			final Range rng = Ranges.range(sheet, rowIndex,	colIdx);
			rng.delete(Range.SHIFT_UP);
		}
	}

	public static void shiftEntireColumnRight(Sheet sheet, int rowIndex, int colIndex) {
		
		int tRow = sheet.getFirstRowNum();
		int bRow = sheet.getPhysicalNumberOfRows();
		for (int rowIdx = tRow; rowIdx < bRow; rowIdx++) {
			final Range rng = Ranges.range(sheet, rowIdx, colIndex);
			rng.insert(Range.SHIFT_RIGHT, Range.FORMAT_RIGHTBELOW);
		}		
	}
	
	public static void shiftEntireColumnLeft(Sheet sheet, int rowIndex, int colIndex) {
		
		int tRow = sheet.getFirstRowNum();
		int bRow = sheet.getPhysicalNumberOfRows();
		for (int rowIdx = tRow; rowIdx < bRow; rowIdx++) {
			final Range rng = Ranges.range(sheet, rowIdx, colIndex);
			rng.delete(Range.SHIFT_LEFT);
		}	
	}
}