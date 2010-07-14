/* Range.java

	Purpose:
		
	Description:
		
	History:
		Mar 10, 2010 2:35:24 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model;

import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Sheet;
import org.zkoss.zss.engine.Ref;

/**
 * Range that represents a cell, a row, a column, or selection of cells containing one or 
 * more contiguous blocks of cells, or a 3-D blocks of cells.
 * 
 * @author henrichen
 */
public interface Range {
	//shift of #insert
	public final static int SHIFT_DEFAULT = 0;
	public final static int SHIFT_RIGHT = 1;
	public final static int SHIFT_DOWN = 2;
	public final static int SHIFT_LEFT = 1;
	public final static int SHIFT_UP = 2;
	
	//copyOrigin of #insert
	public final static int FORMAT_LEFTABOVE = HSSFSheet.FORMAT_LEFTABOVE;
	public final static int FORMAT_RIGHTBELOW = HSSFSheet.FORMAT_RIGHTBELOW;
	
	/**
	 * Returns rich text string of this Range.
	 * @return rich text string of this Range.
	 */
	public RichTextString getText();
	
	/**
	 * Returns formatted text + text color of this Range.
	 * @return formatted text + text color of this Range.
	 */
	public FormatText getFormatText();
	
	/**
	 * Returns the hyperlink of this Range.
	 * @return hyperlink of this Range
	 */
	public Hyperlink getHyperlink();
	
	/**
	 * Return the rich edit text of this Range.
	 * @return the rich edit text of this Range.
	 */
	public RichTextString getRichEditText();
	
	/**
	 * Set {@link RichTextString} as input by the end user.
	 * @param txt the RichTextString object
	 */
	public void setRichEditText(RichTextString txt);
	
	/**
	 * Return the edit text of this Range.
	 * @return the edit text of this Range.
	 */
	public String getEditText();
	
	/**
	 * Set plain text as input by the end user.
	 * @param txt the string input by the end user.
	 */
	public void setEditText(String txt);

	/**
	 * Copy data from this range to the specified destination range.
	 * @param dstRange the destination range.
	 */
	public void copy(Range dstRange);
	
	/**
	 * Pastes a Range from the Clipboard into this range.
	 * @param pasteType the part of the range to be pasted.
	 * @param operation the paste operation
	 * @param SkipBlanks true to not have blank cells in the ranage on the Clipboard pasted into this range; default false.
	 * @param transpose true to transpose rows and columns when pasting to this range; default false.
	 */
	public void pasteSpecial(int pasteType, int operation, boolean SkipBlanks, boolean transpose);
	
	/**
	 * Insert this Range. 
	 * @param shift shiftDown or shiftToRight
	 * @param copyOrigin from where to copy the format to the insert area(FORMAT_LEFTABOVE/FORMAT_RIGHTBELOW)
	 */
	public void insert(int shift, int copyOrigin);
	
	/**
	 * Delete this Range. 
	 * @param shift shiftUp or shiftToLeft
	 */
	public void delete(int shift);

	/**
	 * Sort this Range per the specified parameters
	 * @param rng1 key1 for sorting
	 * @param desc1 true to do descending sort; false to do asceding sort for key1. 
	 * @param rng2 key2 for sorting
	 * @param type PivotTable sorting type(byLabel or byValue); not implemented yet
	 * @param desc2 true to do descending sort; false to do asceding sort for key2.
	 * @param rng3 key3 for sorting
	 * @param desc3 true to do descending sort; false to do asceding sort for key3.
	 * @param header whether sort range includes header
	 * @param orderCustom index of custom order list; not implmented yet 
	 * @param matchCase true to match the string cases; false to ingore string cases
	 * @param sortByRows true to sort by rows(change columns orders); false to sort by columns(change row orders). 
	 * @param sortMethod special sorting method
	 * @param dataOption1 see numeric String as number or not for key1.
	 * @param dataOption2 see numeric String as number or not for key2.
	 * @param dataOption3 see numeric String as number or not for key3.
	 */
	public void sort(Range rng1, boolean desc1, Range rng2, int type, boolean desc2, Range rng3, boolean desc3, int header, int orderCustom,
			boolean matchCase, boolean sortByRows, int sortMethod, int dataOption1, int dataOption2, int dataOption3);

	/**
	 * Merge this range into a merged cell.
	 * @param across true to merge cells in each row; default to false.
	 */
	public void merge(boolean across);

	/**
	 * Un-merge merged cell in this range area to separated cells.
	 */
	public void unMerge();
	
	/**
	 * Adds/Remove border around this range.
	 */
	public void borderAround(BorderStyle lineStyle, String color);

	/**
	 * Adds/Remove border of all cell within this range per the specified border index.
	 */
	public void setBorders(short borderIndex, BorderStyle lineStyle, String color);

	/**
	 * Move this range to a new place as specified by nRow(negative value to move up; 
	 * positive value to move down) and nCol(negative value to move left; positive value to move right)
	 * @param nRow how many rows to move this range
	 * @param nCol how many columns to move this range
	 */
	public void move(int nRow, int nCol);
	
	/**
	 * Sets column width in unit of 1/256 character width of the default font.
	 * @param char256 new column width in unit of 1/256 character width of the default font.
	 */
	public void setColumnWidth(int char256);
	
	/**
	 * Sets row height in points.
	 * @param points new row height in points.
	 */
	public void setRowHeight(int points);

	/**
	 * Get 1st sheet of this range.
	 * @return 1st sheet of this range.
	 */
	public Sheet getFirstSheet();
	
	/**
	 * Get last sheet of this range.
	 * @return last sheet of this range.
	 */
	public Sheet getLastSheet();
	
	/**
	 * Return collection of individual references area of this Range.
	 * @return collection of individual references area of this Range.
	 */
	public List<Ref> getRefs();
	
	/**
	 * Return the range that contains the cell specified in row, col (relative to this Range).
	 * @param row row index relative to this Range
	 * @param col column index relative to this Range
	 * @return the range that contains the cell specified in row, col (relative to this Range).
	 */
	public Range getCells(int row, int col);
}
