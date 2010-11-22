/* Range.java

	Purpose:
		
	Description:
		
	History:
		Mar 10, 2010 2:35:24 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model;

import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zss.model.impl.BookHelper;

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
	public final static int FORMAT_LEFTABOVE = 0;
	public final static int FORMAT_RIGHTBELOW = 1;
	public final static int FORMAT_NONE = -1;
	
	//pasteType of #paste
	public final static int PASTE_ALL = BookHelper.INNERPASTE_FORMATS + BookHelper.INNERPASTE_VALUES_AND_FORMULAS + BookHelper.INNERPASTE_COMMENTS + BookHelper.INNERPASTE_VALIDATION; 
	public final static int PASTE_ALL_EXCEPT_BORDERS = PASTE_ALL - BookHelper.INNERPASTE_BORDERS;
	public final static int PASTE_COLUMN_WIDTHS = BookHelper.INNERPASTE_COLUMN_WIDTHS;
	public final static int PASTE_COMMENTS = BookHelper.INNERPASTE_COMMENTS;
	public final static int PASTE_FORMATS = BookHelper.INNERPASTE_FORMATS; //all formats
	public final static int PASTE_FORMULAS = BookHelper.INNERPASTE_VALUES_AND_FORMULAS; //include values and formulas
	public final static int PASTE_FORMULAS_AND_NUMBER_FORMATS = PASTE_FORMULAS + BookHelper.INNERPASTE_NUMBER_FORMATS;
	public final static int PASTE_VALIDATAION = BookHelper.INNERPASTE_VALIDATION;
	public final static int PASTE_VALUES = BookHelper.INNERPASTE_VALUES;
	public final static int PASTE_VALUES_AND_NUMBER_FORMATS = PASTE_VALUES + BookHelper.INNERPASTE_NUMBER_FORMATS;
	
	//pasteOp of #paste
	public final static int PASTEOP_ADD = BookHelper.PASTEOP_ADD;
	public final static int PASTEOP_SUB = BookHelper.PASTEOP_SUB;
	public final static int PASTEOP_MUL = BookHelper.PASTEOP_MUL;
	public final static int PASTEOP_DIV = BookHelper.PASTEOP_DIV;
	public final static int PASTEOP_NONE = BookHelper.PASTEOP_NONE;
	
	//fillType of #autoFill
	public final static int FILL_COPY = BookHelper.FILL_COPY;
	public final static int FILL_DAYS = BookHelper.FILL_DAYS;
	public final static int FILL_DEFAULT = BookHelper.FILL_DEFAULT;
	public final static int FILL_FORMATS = BookHelper.FILL_FORMATS;
	public final static int FILL_MONTHS = BookHelper.FILL_MONTHS;
	public final static int FILL_SERIES = BookHelper.FILL_SERIES;
	public final static int FILL_VALUES = BookHelper.FILL_VALUES;
	public final static int FILL_WEEKDAYS = BookHelper.FILL_WEEKDAYS;
	public final static int FILL_YEARS = BookHelper.FILL_YEARS;
	public final static int FILL_GROWTH_TREND = BookHelper.FILL_GROWTH_TREND;
	public final static int FILL_LINER_TREND = BookHelper.FILL_LINER_TREND;
	
	//filterOp of #autoFilter
	public final static int FILTEROP_AND = BookHelper.FILTEROP_AND;
	public final static int FILTEROP_BOTTOM10 = BookHelper.FILTEROP_BOTTOM10;
	public final static int FILTEROP_BOTOOM10PERCENT = BookHelper.FILTEROP_BOTOOM10PERCENT;
	public final static int FILTEROP_OR = BookHelper.FILTEROP_OR;
	public final static int FILTEROP_TOP10 = BookHelper.FILTEROP_TOP10;
	public final static int FILTEROP_TOP10PERCENT = BookHelper.FILTEROP_TOP10PERCENT;
		
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
	 * @return the real destination range.
	 */
	public Range copy(Range dstRange);
	
	/**
	 * Pastes a Range from the Clipboard into this range.
	 * @param pasteType the part of the range to be pasted.
	 * @param operation the paste operation
	 * @param SkipBlanks true to not have blank cells in the ranage on the Clipboard pasted into this range; default false.
	 * @param transpose true to transpose rows and columns when pasting to this range; default false.
	 * @return real destination range that was pasted into.
	 */
	public Range pasteSpecial(int pasteType, int operation, boolean SkipBlanks, boolean transpose);
	
	/**
	 * Pastes to a destination Range from this range.
	 * @param dstRange the destination range to be pasted into.
	 * @param pasteType the part of the range to be pasted.
	 * @param operation the paste operation
	 * @param SkipBlanks true to not have blank cells in the ranage on the Clipboard pasted into this range; default false.
	 * @param transpose true to transpose rows and columns when pasting to this range; default false.
	 * @return real destination range that was pasted into.
	 */
	public Range pasteSpecial(Range dstRange, int pasteType, int pasteOp, boolean skipBlanks, boolean transpose);
	
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
	 * @param desc1 true to do descending sort; false to do ascending sort for key1. 
	 * @param rng2 key2 for sorting
	 * @param type PivotTable sorting type(byLabel or byValue); not implemented yet
	 * @param desc2 true to do descending sort; false to do ascending sort for key2.
	 * @param rng3 key3 for sorting
	 * @param desc3 true to do descending sort; false to do ascending sort for key3.
	 * @param header whether sort range includes header
	 * @param orderCustom index of custom order list; not implmented yet 
	 * @param matchCase true to match the string cases; false to ignore string cases
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
	 * Returns associate sheet of this range.
	 * @return associate sheet of this range.
	 */
	public Sheet getSheet();
	
	/**
	 * Return the range that contains the cell specified in row, col (relative to this Range).
	 * @param row row index relative to this Range(note that it is 0-based)
	 * @param col column index relative to this Range(note that it is 0-based)
	 * @return the range that contains the cell specified in row, col (relative to this Range).
	 */
	public Range getCells(int row, int col);
	
	/**
	 * Sets a Style object to this Range.
	 * @param style the style object
	 */
	public void setStyle(CellStyle style);
	
	/**
	 * Perform an auto fill on the specified destination Range. Note the given destination Range
	 * must include this source Range.
	 * @param dstRange destination range to do the auto fill. Note the given destination Range must include this source Range
	 * @param fillType the fillType
	 */
	public void autoFill(Range dstRange, int fillType);
	
	/**
	 * Clears the data from this Range.
	 */
	public void clearContents();
	
	/**
	 * Fills down from the top cells of this Range to the rest of this Range.
	 */
	public void fillDown();
	
	/**
	 * Fills left from the rightmost cells of this Range to the rest of this Range.
	 */
	public void fillLeft();
	
	/**
	 * Fills right from the leftmost cells of this Range to the rest of this Range.
	 */
	public void fillRight();

	/**
	 * Fills up from the bottom cells of this Range to the rest of this Range.
	 */
	public void fillUp();
	
	/**
	 * Filters a list specified by this Range.
	 * @param field offset of the field on which you want to base the filter on (1-based; i.e. leftmost column in this range is field 1).
	 * @param criteria1 "=" to find blank fields, "<>" to find non-blank fields. If null, means ALL. If filterOp == Range#FILTEROP_TOP10, 
	 * then this shall specifies the number of items (e.g. "10"). 
	 * @param filterOp see Range#FILTEROP_xxx. Use FILTEROP_AND and FILTEROP_OR with criteria1 and criterial2 to construct compound criteria.
	 * @param criteria2 2nd criteria; used with criteria1 and filterOP to construct compound criteria.
	 * @param visibleDropDown true to show the autoFilter drop-down arrow for the filtered field; false to hide the autoFilter drop-down arrow.
	 */
//TODO UNTIL POI support reading/writing the autoFilter record	
//	public void autoFilter(int field, String criteria1, int filterOp, String criteria2, boolean visibleDropDown);
	
	/**
	 * Sets whether this rows or columns are hidden(useful only if this Range cover entire column or entire row)
	 * @param hidden true to hide this rows or columns
	 */
	public void setHidden(boolean hidden);
	
	/**
	 * Sets whether show the gridlines of the sheets in this Range.
	 * @param show true to show the gridlines; false to not show the gridlines. 
	 */
	public void setDisplayGridlines(boolean show);
	
	/**
	 * Sets the hyperlink of this Range
	 * @param linkType the type of target to link. One of the {@link #Hyperlink.LINK_URL}, 
	 * {@link #Hyperlink.LINK_DOCUMENT}, {@link #Hyperlink.LINK_EMAIL}, {@link #LINK_FILE}
	 * @param address the address
	 * @param display the text to display link
	 */
	public void setHyperlink(int linkType, String address, String display);
	
	/**
	 * Returns an {@link Areas} which is a collection of each single selected area(also Range) of this multiple-selected Range. 
	 * If this Range is a single selected Range, this method return the Areas which contains only this Range itself.
	 * @return
	 */
	public Areas getAreas();
	
	/**
	 * Returns a {@link Range} that represent all columns of the 1st selected area of this Range. Note that only the 1st selected area is considered if this Range is a multiple-selected Range. 
	 * @return a {@link Range} that represent all columns of this Range.
	 */
	public Range getColumns();
	
	/**
	 * Returns a {@link Range} that represent all rows of the 1st selected area of this Range. Note that only the 1st selected area is considered if this Range is a multiple-selected Range. 
	 * @return a {@link Range} that represent all rows of this Range.
	 */
	public Range getRows();

	/**
	 * Returns a {@link Range} that represent all dependents of the left-top cell of the 1st selected area of this Range. 
	 * Note that only the left-top cell of the 1st selected area is considered if this Range is a multiple-selected Range.
	 * This could be multiple-selected Range if there are more than one dependent. 
	 * @return a {@link Range} that represent all dependents of the left-top cell of the 1st selected area of this Range.
	 */
	public Range getDependents();
	
	/**
	 * Returns a {@link Range} that represent all direct dependents of the left-top cell of the 1st selected area of this Range. 
	 * Note that only the left-top cell of the 1st selected area is considered if this Range is a multiple-selected Range. 
	 * This method could return multiple-selected Range if there are more than one dependent. 
	 * @return a {@link Range} that represent all direct dependents of the left-top cell of the 1st selected area of this Range.
	 */
	public Range getDirectDependents();
	
	/** 
	 * Returns a {@link Range} that represent all precedents of the left-top cell of the 1st selected area of this Range. 
	 * Note that only the left-top cell of the 1st selected area is considered if this Range is a multiple-selected Range.
	 * This method could return multiple-selected Range if there are more than one precedent. 
	 * @return a {@link Range} that represent all precedents of the left-top cell of the 1st selected area of this Range.
	 */
	public Range getPrecedents();
	
	/**
	 * Returns a {@link Range} that represent all direct precedents of the left-top cell of the 1st selected area of this Range. 
	 * Note that only the left-top cell of the 1st selected area is considered if this Range is a multiple-selected Range. 
	 * This method could return multiple-selected Range if there are more than one precedent. 
	 * @return a {@link Range} that represent all direct precedents of the left-top cell of the 1st selected area of this Range.
	 */
	public Range getDirectPrecedents();
	
	/**
	 * Returns the number of the 1st row of the 1st area in this Range(0-based; i.e. row1 return 0)
	 * @return the number of the 1st row of the 1st area in this Range(0-based; i.e. row1 return 0)
	 */
	public int getRow();
	
	/**
	 * Returns the number of the 1st column of the 1st area in this Range(0-based; i.e. Column A return 0)
	 * @return the number of the 1st column of the 1st area in this Range(0-based; i.e. Column A return 0)
	 */
	public int getColumn();
	
	/**
	 * Returns the number of the last row of the 1st area in this Range(0-based; i.e. row1 return 0)
	 * @return the number of the last row of the 1st area in this Range(0-based; i.e. row1 return 0)
	 */
	public int getLastRow();
	
	/**
	 * Returns the number of the last column of the 1st area in this Range(0-based; i.e. Column A return 0)
	 * @return the number of the last column of the 1st area in this Range(0-based; i.e. Column A return 0)
	 */
	public int getLastColumn();
	
	/**
	 * Returns the number of contained objects in this Range.
	 * @return the number of contained objects in this Range.
	 */
	public long getCount();

	/**
	 * Set value into the specified Range.
	 */
	public void setValue(Object value);
	
	/**
	 * Returns value from the specified Range.
	 * @return
	 */
	public Object getValue();
	
	/**
	 * Returns a {@link Range} that represents a range that offset from this Range. 
	 * @param rowOffset positive means downward; 0 means don't change row; negative means upward.
	 * @param colOffset positive means rightward; 0 means don't change column; negative means leftward.
	 * @return
	 */
	public Range getOffset(int rowOffset, int colOffset);
}
