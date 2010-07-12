/* Utils.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 20, 2008 12:40:01 PM     2008, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

import java.awt.font.FontRenderContext;
import java.awt.font.TextAttribute;
import java.awt.font.TextLayout;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.text.AttributedString;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.zkoss.image.AImage;
import org.zkoss.util.logging.Log;
import org.zkoss.zk.ui.Execution;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.engine.RefBook;
import org.zkoss.zss.engine.RefSheet;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.FormatText;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.RangeImpl;
import org.zkoss.zss.ui.Spreadsheet;

/**
 * Utility class for {@link Spreadsheet}.
 * @author Dennis.Chen
 *
 */
public class Utils {
	private static final Log log = Log.lookup(Utils.class);
	
	/**
	 * Format and escape a {@link Hyperlink} to HTML &lt;a> string.
	 * @param sheet the sheet with the RichTextString 
	 * @param hlink the Hyperlink
	 * @return the HTML &lt;a> format string
	 */
	public static String formatHyperlink(Sheet sheet, Hyperlink hlink, boolean wrap) {
		if (hlink == null) {
			return "";
		}
		final List<int[]> indexes = new ArrayList<int[]>(2);
		String text = BookHelper.formatHyperlink((Book)sheet.getWorkbook(), hlink, indexes);
		return Utils.escapeCellText(text, wrap, wrap, indexes);
	}
	/**
	 * Format and escape a {@link RichTextString} to HTML &lt;span> string.
	 * @param sheet the sheet with the RichTextString 
	 * @param rstr the RichTextString
	 * @param wrap whether wrap the string if see "\n".
	 * @return the HTML &lt;span> format string
	 */
	public static String formatRichTextString(Sheet sheet, RichTextString rstr, boolean wrap) {
		final List<int[]> indexes = new ArrayList<int[]>(rstr.numFormattingRuns()+1);
		String text = BookHelper.formatRichText((Book)sheet.getWorkbook(), rstr, indexes);
		return Utils.escapeCellText(text, wrap, wrap, indexes);
	}

	/**
	 * Escape character that has special meaning in HTML such as &lt;, &amp;, etc..
	 * @param text the text
	 * @param wrap whether to allow wrap
	 * @param multiline whether to show multiple line
	 * @return the HTML 
	 */
	public static String escapeCellText(String text, boolean wrap, boolean multiline) {
		final StringBuffer out = new StringBuffer();
		for (int j = 0, tl = text.length(); j < tl; ++j) {
			char cc = text.charAt(j);
			switch (cc) {
			case '&': out.append("&amp;"); break;
			case '<': out.append("&lt;"); break;
			case '>': out.append("&gt;"); break;
			case ' ': out.append(wrap?" ":"&nbsp;"); break;
			case '\n':
				if (multiline) {
					out.append("<br/>");
					break;
				}
			default:
				out.append(cc);
			}
		}
		return out.toString();
	}
	
	//escape character that has special meaning in HTML such as &lt;, &amp;, etc..
	//runs is a index pair to the text string that needs to be escaped
	private static String escapeCellText(String text,boolean wrap,boolean multiline, List<int[]> runs){
		
		StringBuffer out = new StringBuffer();
		if (text!=null){
			int j = 0;
			for (int[] run : runs) {
				for (int tl = run[0]; j < tl; ++j) {
					char cc = text.charAt(j);
					out.append(cc);
				}
				for (int tl = run[1]; j < tl; ++j) {
					char cc = text.charAt(j);
					switch (cc) {
					case '&': out.append("&amp;"); break;
					case '<': out.append("&lt;"); break;
					case '>': out.append("&gt;"); break;
					case ' ': out.append(wrap?" ":"&nbsp;"); break;
					case '\n':
						if (multiline) {
							out.append("<br/>");
							break;
						}
					default:
						out.append(cc);
					}
				}
			}
			for (int tl = text.length(); j < tl; ++j) {
				char cc = text.charAt(j);
				out.append(cc);
			}
		}
		return out.toString();
	}
	
	//Used for special id generator so UI component can find the sheet easily.
	public static String nextUpdateId(){
		Execution ex = Executions.getCurrent();
		Integer count;
		synchronized (ex) {
			count = (Integer) ex.getAttribute("_zssmseq");
			if (count == null) {
				count = new Integer(0);
			} else {
				count = new Integer(count.intValue() + 1);
			}
			ex.setAttribute("_zssmseq", count);
		}
		return count.toString();
	}
	
	/**
	 * Return or create if not exist the {@link Cell} per the given sheet, row index, and column index. 
	 * @param sheet the sheet to get cell from
	 * @param rowIndex the row index of the cell
	 * @param colIndex the column index of the cell
	 * @return or create if not exist the {@link Cell} per the given sheet, row index, and column index. 
	 */
	public static Cell getOrCreateCell(Sheet sheet,int rowIndex, int colIndex){
		Row row = getOrCreateRow(sheet, rowIndex);
		Cell cell = row.getCell(colIndex);
		if (cell == null) {
			cell = row.createCell(colIndex);
		}
		return cell;
	}
	
	public static Row getOrCreateRow(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null) {
			row = sheet.createRow(rowIndex);
		}
		return row;
	}
	
	/**
	 * Return the {@link Cell} per the given sheet, row index, and column index; return null if cell not exists. 
	 * @param sheet the sheet to get cell from
	 * @param rowIndex the row index of the cell
	 * @param colIndex the column index of the cell
	 * @return the {@link Cell} per the given sheet, row index, and column index; return null if cell not exists. 
	 */
	public static Cell getCell(Sheet sheet,int rowIndex, int colIndex){
		final Row row = sheet.getRow(rowIndex);
		return row != null ? row.getCell(colIndex) : null; 
	}
	
	public static void copyCell(Sheet srcSheet, int srcRow, int srcCol, Sheet dstSheet, int dstRow, int dstCol) {
		final Range srcRange = getRange(srcSheet, srcRow, srcCol);
		final Range dstRange = getRange(dstSheet, dstRow, dstCol);
		srcRange.copy(dstRange);
	}
	
	public static void copyCell(Cell cell, Sheet dstSheet, int dstRow, int dstCol) {
		copyCell(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex(), dstSheet, dstRow, dstCol);
	}
	
	public static void insertRows(Sheet sheet, int startRow, int endRow) {
		final Book book = (Book)sheet.getWorkbook();
		final Range rng = getRange(sheet, startRow, 0, endRow, book.getSpreadsheetVersion().getLastColumnIndex());
		rng.insert(Range.SHIFT_DEFAULT, 0);
	}

	public static void deleteRows(Sheet sheet, int startRow, int endRow) {
		final Book book = (Book)sheet.getWorkbook();
		final Range rng = getRange(sheet, startRow, 0, endRow, book.getSpreadsheetVersion().getLastColumnIndex());
		rng.delete(Range.SHIFT_DEFAULT);
	}

	public static void insertColumns(Sheet sheet, int startCol, int endCol) {
		final Book book = (Book)sheet.getWorkbook();
		final Range rng = getRange(sheet, 0, startCol, book.getSpreadsheetVersion().getLastRowIndex(), endCol);
		rng.insert(Range.SHIFT_DEFAULT, 0);
	}

	public static void deleteColumns(Sheet sheet, int startCol, int endCol) {
		final Book book = (Book)sheet.getWorkbook();
		final Range rng = getRange(sheet, 0, startCol, book.getSpreadsheetVersion().getLastRowIndex(), endCol);
		rng.delete(Range.SHIFT_DEFAULT);
	}

	public static void setCellValue(Sheet sheet, int rowIndex, int colIndex, String value) {
		final Cell cell = getOrCreateCell(sheet, rowIndex, colIndex);
		cell.setCellValue(value);
	}
	public static void setCellValue(Sheet sheet, int rowIndex, int colIndex, double value) {
		final Cell cell = getOrCreateCell(sheet, rowIndex, colIndex);
		cell.setCellValue(value);
	}
	public static void setCellValue(Sheet sheet, int rowIndex, int colIndex, boolean value) {
		final Cell cell = getOrCreateCell(sheet, rowIndex, colIndex);
		cell.setCellValue(value);
	}
	public static void setCellValue(Sheet sheet, int rowIndex, int colIndex, Date value) {
		final Cell cell = getOrCreateCell(sheet, rowIndex, colIndex);
		cell.setCellValue(value);
	}
	public static void setCellValue(Sheet sheet, int rowIndex, int colIndex, int value) {
		final Cell cell = getOrCreateCell(sheet, rowIndex, colIndex);
		cell.setCellValue(value);
	}
	
	/**
	 * Return a cell {@link Range} per the given sheet, row index, and column index.
	 * @param sheet the sheet to get the {@link Range} from. 
	 * @param row the row index of the cell
	 * @param col the column index of the cell
	 * @return a cell {@link Range} per the given sheet, row index, and column index.
	 */
	public static Range getRange(Sheet sheet, int row, int col) {
		return new RangeImpl(row, col, sheet, sheet);
	}

	/**
	 * Returns a area {@link Range} per the given sheet, top row index, left column index, 
	 * bottom row index, and right column index.
	 * @param sheet the sheet to get the {@link Range} from. 
	 * @param tRow the top row index of the area
	 * @param lCol the left column index of the area
	 * @param bRow the bottom row index of the area
	 * @param rCol the right column index of the area
	 * @return a area {@link Range} per the given sheet, top row index, left column index, 
	 * bottom row index, and right column index.
	 */
	public static Range getRange(Sheet sheet, int tRow, int lCol, int bRow, int rCol) {
		return new RangeImpl(tRow, lCol, bRow, rCol, sheet, sheet);
	}
	
	/**
	 * Return a 3D cell {@link Range} per the given sheet, row index, and column index.
	 * @param firstSheet the sheet to get the {@link Range} from. 
	 * @param lastSheet the sheet to get the {@link Range} from. 
	 * @param row the row index of the cell
	 * @param col the column index of the cell
	 * @return a cell {@link Range} per the given sheet, row index, and column index.
	 */
	public static Range getRange(Sheet firstSheet, Sheet lastSheet, int row, int col) {
		return new RangeImpl(row, col, firstSheet, lastSheet);
	}
	
	/**
	 * Return a 3D area {@link Range} per the given sheet, top row index, left column index, 
	 * bottom row index, and right column index.
	 * @param firstSheet the sheet to get the {@link Range} from. 
	 * @param lastSheet the sheet to get the {@link Range} from. 
	 * @param tRow the top row index of the area
	 * @param lCol the left column index of the area
	 * @param bRow the bottom row index of the area
	 * @param rCol the right column index of the area
	 * @return a 3D area {@link Range} per the given sheet, top row index, left column index, 
	 * bottom row index, and right column index.
	 */
	public static Range getRange(Sheet firstSheet, Sheet lastSheet, int tRow, int lCol, int bRow, int rCol) {
		return new RangeImpl(tRow, lCol, bRow, rCol, firstSheet, lastSheet);
	} 
	
	/**
	 * Returns the id of the spcified {@link Sheet}. 
	 * to identify a {@link Sheet})
	 * @param sheet the sheet
	 * @return the id of the spcified {@link Sheet}.
	 */
	public static String getSheetId(Sheet sheet){
		return ""+sheet.hashCode();
	}
	
	/**
	 * Returns the {@link Sheet} of the specified id; null if id not exists.
	 * @param book the book the contains the {@link Sheet}
	 * @param id the sheet id
	 * @return the {@link Sheet} of the specified id; null if id not exists.
	 */
	public static Sheet getSheetById(Book book, String id) {
		int count = book.getNumberOfSheets();
		for(int j = 0; j < count; ++j) {
			Sheet sheet = book.getSheetAt(j);
			if (id.equals(getSheetId(sheet))) {
				return sheet;
			}
		}
		return null;
	}
	
	/**
	 * Returns whether the {@link Spreadsheet} title is in index mode
	 * @param ss the {@link Spreadsheet} 
	 * @return whether the {@link Spreadsheet} title is in index mode
	 */
	public static boolean isTitleIndexMode(Spreadsheet ss){
		return "index".equals(ss.getAttribute("zss_titlemode"));
	}
	
	/**
	 * Returns the associated model sheet({@link Sheet}) per the given the reference sheet({@link RefSheet}). 
	 * @param book the model book
	 * @param refSheet the reference sheet
	 * @return the associated model sheet({@link Sheet}) per the given the reference sheet({@link RefSheet}).
	 */
	public static Sheet getSheetByRefSheet(Book book, RefSheet refSheet) {
		return book.getSheet(refSheet.getSheetName());
	}

	/**
	 * Returns the {@link Hyperlink} to be shown on the specified cell.
	 * @param cell the cell
	 * @return the {@link Hyperlink} to be shown on the specified cell.
	 */
	public static Hyperlink getHyperlink(Cell cell) {
		Range range = getRange(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex());
		return range.getHyperlink();
	}
	/**
	 * Returns the {@link RichTextString} to be shown on the specified cell. 
	 * @param cell the cell
	 * @return the {@link RichTextString} to be shown on the specified cell.
	 */
	public static RichTextString getText(Cell cell) {
		Range range = getRange(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex());
		return range.getText();
	}
	
	public static FormatText getFormatText(Cell cell) {
		Range range = getRange(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex());
		return range.getFormatText();
	}
	
	/**
	 * Returns the text for editing on the specified cell. 
	 * @param cell the cell
	 * @return the text for editing on the specified cell.
	 */
	public static String getEditText(Cell cell) {
		final RichTextString rstr = cell == null ? null : getRichEditText(cell);
		return rstr != null ? rstr.getString() : "";
	}
	
	/**
	 * Set user input text into the cell at specified sheet, row index, and column index.
	 * @param sheet the sheet where the cell resides
	 * @param row the row index
	 * @param col the column index
	 * @param value the user input text
	 */
	public static void setEditText(Sheet sheet, int row, int col, String value) {
		//System.out.println("setEditText val: " + value);
		final Cell cell = getOrCreateCell(sheet, row, col);
		setEditText(cell, value);
	}
	
	/**
	 * Sets the user input text into the cell.
	 * @param cell the cell.
	 * @param txt the user input text.
	 */
	public static void setEditText(Cell cell, String txt) {
		final Range range = getRange(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex());
		range.setEditText(txt);
	}

	/**
	 * Returns the {@link RichTextString} for editing on the cell.
	 * @param cell the cell
	 * @return the {@link RichTextString} for editing on the cell.
	 */
	public static RichTextString getRichEditText(Cell cell) {
		final Range range = getRange(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex());
		return range.getRichEditText();
	}
	
	/**
	 * Sets the user input {@link RichTextString} into the cell.
	 * @param cell the cell
	 * @param value the user input {@link RichTextString} to be set into the cell.
	 */
	public static void setRichEditText(Cell cell, RichTextString value) {
		final Range range = getRange(cell.getSheet(), cell.getRowIndex(), cell.getColumnIndex());
		range.setRichEditText(value);
	}
	
	public static AImage getAImage(Sheet sheet, PictureData picdata, int pictureIndex, String name) {
		try {
			return new AImage(name, picdata.getData());
		} catch (IOException e) {
			// ignore
		}
		return null;
	}
	
	public static void mergeCells(Sheet sheet, int tRow, int lCol, int bRow, int rCol, boolean across) {
		Range rng = getRange(sheet, tRow, lCol, bRow, rCol);
		rng.merge(across);
	}
	
	public static void unmergeCells(Sheet sheet, int tRow, int lCol, int bRow, int rCol) {
		Range rng = getRange(sheet, tRow, lCol, bRow, rCol);
		rng.unMerge();
	}
	
	public static int getWidthInPx(Sheet zkSheet, ClientAnchor anchor, int charWidth) {
	    final int l = anchor.getCol1();
	    final int lfrc = anchor.getDx1();
	    
	    //first column
	    final int lw = getWidthAny(zkSheet,l, charWidth);
	    
	    final int wFirst = lfrc >= 1024 ? 0 : (lw - lw * lfrc / 1024);  
	    
	    //last column
	    final int r = anchor.getCol2();
	    int wLast = 0;
	    if (l != r) {
		    final int rfrc = anchor.getDx2();
	    	final int rw = getWidthAny(zkSheet,r, charWidth);
	    	wLast = rw * rfrc / 1024;  
	    }
	    
	    //in between
	    int width = wFirst + wLast;
	    for (int j = l+1; j < r; ++j) {
	    	width += getWidthAny(zkSheet,j, charWidth);
	    }
	    
	    return width;
	}
	
	public static int getHeightInPx(Sheet zkSheet, ClientAnchor anchor) {
	    final int t = anchor.getRow1();
	    final int tfrc = anchor.getDy1();
	    
	    //first row
	    final int th = getHeightAny(zkSheet,t);
	    final int hFirst = tfrc >= 256 ? 0 : (th - th * tfrc / 256);  
	    
	    //last row
	    final int b = anchor.getRow2();
	    int hLast = 0;
	    if (t != b) {
		    final int bfrc = anchor.getDy2();
	    	final int bh = getHeightAny(zkSheet,b);
	    	hLast = bh * bfrc / 256;  
	    }
	    
	    //in between
	    int height = hFirst + hLast;
	    for (int j = t+1; j < b; ++j) {
	    	height += getHeightAny(zkSheet,j);
	    }
	    
	    return height;
	}
	
	public static int getTopFraction(Sheet zkSheet, ClientAnchor anchor) {
	    final int t = anchor.getRow1();
	    final int tfrc = anchor.getDy1();
	    
	    //first row
	    final int th = getHeightAny(zkSheet,t);
	    return tfrc >= 256 ? th : (th * tfrc / 256);  
	}
	
	public static int getLeftFraction(Sheet zkSheet, ClientAnchor anchor, int charWidth) {
	    final int l = anchor.getCol1();
	    final int lfrc = anchor.getDx1();
	    
	    //first column
	    final int lw = getWidthAny(zkSheet,l, charWidth);
	    return lfrc >= 1024 ? lw : (lw * lfrc / 1024);  
	}
	
	private static int getWidthAny(Sheet zkSheet,int col, int charWidth){
		int w = zkSheet.getColumnWidth(col);
		if (w == zkSheet.getDefaultColumnWidth() * 256) { //default column width
			return Utils.defaultColumnWidthToPx(w / 256, charWidth);
		}
		return fileChar256ToPx(w, charWidth);
	}
	
	private static int getHeightAny(Sheet zkSheet,int row){
		int h = zkSheet.getRow(row).getHeight();
		if(h==-1) h = zkSheet.getDefaultRowHeight();
		return twipToPx(h);
	}
	
	//calculate the default char width in pixel per the given Font
	public static int calcDefaultCharWidth(java.awt.Font font) {
        /**
         * Excel measures columns in units of 1/256th of a character width
         * but the docs say nothing about what particular character is used.
         * '0' looks to be a good choice. ref. http://support.microsoft.com/kb/214123
         */
        final String defaultString = "0";

        final FontRenderContext frc = new FontRenderContext(null, true, true);
        final AttributedString str = new AttributedString(defaultString);
        copyAttributes(font, str, 0, defaultString.length());
        final TextLayout layout = new TextLayout(str.getIterator(), frc);
        //TODO, don't know how to calculate the charter width per the Font
        final double w = layout.getAdvance();
        final int charWidth = (int) Math.floor(w + 0.5) +  1;
        return charWidth;
	}
	
    private static void copyAttributes(java.awt.Font font, AttributedString str, int startIdx, int endIdx) {
        str.addAttribute(TextAttribute.FAMILY, font.getFontName(), startIdx, endIdx);
        str.addAttribute(TextAttribute.SIZE, new Float(font.getSize2D()), startIdx, endIdx);
        if (font.isBold()) str.addAttribute(TextAttribute.WEIGHT, TextAttribute.WEIGHT_BOLD, startIdx, endIdx);
        if (font.isItalic()) str.addAttribute(TextAttribute.POSTURE, TextAttribute.POSTURE_OBLIQUE, startIdx, endIdx);
    }

	/** convert pixel to point */
	public static int pxToPoint(int px) {
		return px * 72 / 96; //assume 96dpi
	}
	
	/** convert pixel to twip (1/20 point) */
	public static int pxToTwip(int px) {
		return px * 72 * 20 / 96; //assume 96dpi
	}

	/** convert twip (1/20 point) to pixel */
	public static int twipToPx(int twip) {
		return twip * 96 / 72 / 20; //assume 96dpi
	}

	/** convert file 1/256 character width to pixel */
	public static int fileChar256ToPx(int char256, int charWidth) {
		final double w = (double) char256;
		return (int) Math.floor(w * charWidth / 256 + 0.5);
	}
	
	/** convert pixel to file 1/256 character width */
	public static int pxToFileChar256(int px, int charWidth) {
		final double w = (double) px;
		return (int) Math.floor(px * 256 / charWidth + 0.5);
	}
	
	/** convert 1/256 character width to pixel */
	public static int screenChar256ToPx(int char256, int charWidth) {
		final double w = (double) char256;
		return (char256 < 256) ?
			(int) Math.floor(w * (charWidth + 5) / 256 + 0.5):
			(int) Math.floor(w * charWidth  / 256 + 0.5) + 5;
	}
	
	/** Convert pixel to 1/256 character width */
	public static int pxToScreenChar256(int px, int charWidth) {
		return px < (charWidth + 5) ? 
				px * 256 / (charWidth + 5):
				(px - 5) * 256 / charWidth;
	}
	
	/** convert character width to pixel */
	public static int screenChar1ToPx(double w, int charWidth) {
		return w < 1 ?
			(int) Math.floor(w * (charWidth + 5) + 0.5):
			(int) Math.floor(w * charWidth + 0.5) + 5;
	}
	
	/** Convert pixel to character width (for application) */
	public static double pxToScreenChar1(int px, int charWidth) {
		final double w = (double) px;
		return px < (charWidth + 5) ?
			roundTo100th(w / (charWidth + 5)):
			roundTo100th((w - 5) / charWidth);
	}

	/** Convert default columns character width to pixel */ 
	public static int defaultColumnWidthToPx(int columnWidth, int charWidth) {
		final int w = columnWidth * charWidth + 5;
		final int diff = w % 8;
		return w + (diff > 0 ? (8 - diff) : 0);
	}
	
	private static double roundTo100th(double w) {
		return Math.floor(w * 100 + 0.5) / 100;
	}
	
	public static String getColumnTitle(Sheet sheet, int col) {
		return CellReference.convertNumToColString(col);
	}
	
	public static String getRowTitle(Sheet sheet, int row) {
		return ""+(row+1);
	}

	/**
	 * Sets column Width in pixel.
	 * @param sheet
	 * @param col
	 * @param px column width in pixel
	 */
	public static void setColumnWidth(Sheet sheet, int col, int px) {
		final int char256 = Utils.pxToFileChar256(px, ((Book)sheet.getWorkbook()).getDefaultCharWidth());
		final Range rng = Utils.getRange(sheet, -1 , col);
		rng.setColumnWidth(char256);
	}
	
	/**
	 * Sets Row Height in pixel. 
	 * @param sheet
	 * @param row
	 * @param px row height in pixel 
	 */
	public static void setRowHeight(Sheet sheet, int row, int px) {
		final int points = Utils.pxToPoint(px);
		final Range rng = Utils.getRange(sheet, row, -1);
		rng.setRowHeight(points);
	}
	
	public static void moveRows(Sheet sheet, int tRow, int bRow, int nRow) {
		final int maxcol = ((Book)sheet.getWorkbook()).getSpreadsheetVersion().getLastColumnIndex();
		final Range rng = Utils.getRange(sheet, tRow, 0, bRow, maxcol);
		rng.move(nRow, 0);
	}

	public static void moveColumns(Sheet sheet, int lCol, int rCol, int nCol) {
		final int maxrow = ((Book)sheet.getWorkbook()).getSpreadsheetVersion().getLastRowIndex();
		final Range rng = Utils.getRange(sheet, 0, lCol, maxrow, rCol);
		rng.move(0, nCol);
	}
	
	public static void moveCells(Sheet sheet, int tRow, int lCol, int bRow, int rCol, int nRow, int nCol) {
		final Range rng = Utils.getRange(sheet, tRow, lCol, bRow, rCol);
		rng.move(nRow, nCol);
	}
}
