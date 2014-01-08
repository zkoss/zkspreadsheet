/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngapi;

import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.zss.ngmodel.NAutoFilter;
import org.zkoss.zss.ngmodel.NAutoFilter.FilterOp;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NDataValidation;
import org.zkoss.zss.ngmodel.NHyperlink;
import org.zkoss.zss.ngmodel.NHyperlink.HyperlinkType;
import org.zkoss.zss.ngmodel.NSheet;
/**
 * The most useful api to manipulate a book
 *  
 * @author dennis
 * @since 3.5.0
 */
public interface NRange {
	
	public ReadWriteLock getLock();
	
	public enum PasteType{
		ALL,
		ALL_EXCEPT_BORDERS,
		COLUMN_WIDTHS,
		COMMENTS,
		FORMATS/*all formats*/,
		FORMULAS/*include values and formulas*/,
		FORMULAS_AND_NUMBER_FORMATS,
		VALIDATAION,
		VALUES,
		VALUES_AND_NUMBER_FORMATS;
	}
	
	public enum PasteOperation{
		ADD,
		SUB,
		MUL,
		DIV,
		NONE;
	}
	
	public enum ApplyBorderType{
		FULL,
		EDGE_BOTTOM,
		EDGE_RIGHT,
		EDGE_TOP,
		EDGE_LEFT,
		OUTLINE,
		INSIDE,
		INSIDE_HORIZONTAL,
		INSIDE_VERTICAL,
		DIAGONAL,
		DIAGONAL_DOWN,
		DIAGONAL_UP
	}
	
	/** Shift direction of insert api**/
	public enum InsertShift{
		DEFAULT,
		RIGHT,
		DOWN
	}
	/** 
	 * Copy origin format/style of insert
	 **/
	public enum InsertCopyOrigin{
		FORMAT_NONE,
		FORMAT_LEFT_ABOVE,
		FORMAT_RIGHT_BELOW,
	}
	/** Shift direction of delete api**/
	public enum DeleteShift{
		DEFAULT,
		LEFT,
		UP
	}
	
	public enum SortDataOption{
		NORMAL_DEFAULT,
		TEXT_AS_NUMBERS
	}
	
	public enum AutoFilterOperation{
		AND,
		BOTTOM10,
		BOTOOM10PERCENT,
		OR,
		TOP10,
		TOP10PERCENT,
		VALUES
	}
	
	public enum AutoFillType{
		COPY,
		DAYS,
		DEFAULT,
		FORMATS,
		MONTHS,
		SERIES,
		VALUES,
		WEEKDAYS,
		YEARS,
		GROWTH_TREND,
		LINER_TREND
	}
	
//	public NSheet getSheet();
//	public int getRow();
//	public int getColumn();
//	public int getLastRow();
//	public int getLastColumn();
//	
//	public void setEditText(String editText);
//	public String getEditText();
//	
//	public void setValue(Object value);
//	public void clear();
//	public void notifyChange();
//	public boolean isWholeRow();
//	public NRange getRows();
//	public void setRowHeight(int heightPx);
//	public boolean isWholeColumn();
//	public NRange getColumns();
//	public void setColumnWidth(int widthPx);
//	boolean isWholeSheet();
	
	////////////////////////////////////
//	/**
//	 * Returns rich text string of this Range.
//	 * @return rich text string of this Range.
//	 */
//	public RichTextString getText();
//	
//	/**
//	 * Returns formatted text + text color of this Range.
//	 * @return formatted text + text color of this Range.
//	 */
//	public XFormatText getFormatText();
	
	/**
	 * Returns the hyperlink of this Range.
	 * @return hyperlink of this Range
	 */
	public NHyperlink getHyperlink();
	
//	/**
//	 * Return the rich edit text of this Range.
//	 * @return the rich edit text of this Range.
//	 */
//	public RichTextString getRichEditText();
//	
//	/**
//	 * Set {@link RichTextString} as input by the end user.
//	 * @param txt the RichTextString object
//	 */
//	public void setRichEditText(RichTextString txt);
	
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
	 * cut the selected range and paste to destination range.
	 * @param dstRange
	 * @return the real destination range.
	 * @since 3.0.0
	 */
	public NRange copy(NRange dstRange, boolean cut);

	/**
	 * Copy data from this range to the specified destination range.
	 * @param dstRange the destination range.
	 * @return the real destination range.
	 */
	public NRange copy(NRange dstRange);
	
	/**
	 * Pastes to a destination Range from this range.
	 * @param dstRange the destination range to be pasted into.
	 * @param pasteType the part of the range to be pasted.
	 * @param pasteOp the paste operation
	 * @param skipBlanks true to not have blank cells in the ranage to paste into destination Range; default false.
	 * @param transpose true to transpose rows and columns when pasting to this range; default false.
	 * @return real destination range that was pasted into.
	 */
	public NRange pasteSpecial(NRange dstRange, PasteType pasteType, PasteOperation pasteOp, boolean skipBlanks, boolean transpose);
	
	/**
	 * Insert this Range. 
	 * @param shift can be {@link #SHIFT_DEFAULT}, {{@link #SHIFT_DOWN}, or {@link #SHIFT_RIGHT}.
	 * @param copyOrigin from where to copy the format to the insert area({@link #FORMAT_LEFTABOVE} /{@link #FORMAT_RIGHTBELOW})
	 */
	public void insert(InsertShift shift, InsertCopyOrigin copyOrigin);
	
	/**
	 * Delete this Range. 
	 * @param shift can be {@link #SHIFT_DEFAULT}, {{@link #SHIFT_UP}, or {@link #SHIFT_LEFT}.
	 */
	public void delete(DeleteShift shift);

//	/**
//	 * Sort this Range per the specified parameters
//	 * @param rng1 key1 for sorting
//	 * @param desc1 true to do descending sort; false to do ascending sort for key1. 
//	 * @param rng2 key2 for sorting
//	 * @param type PivotTable sorting type(byLabel or byValue); not implemented yet
//	 * @param desc2 true to do descending sort; false to do ascending sort for key2.
//	 * @param rng3 key3 for sorting
//	 * @param desc3 true to do descending sort; false to do ascending sort for key3.
//	 * @param header whether sort range includes header
//	 * @param orderCustom index of custom order list; not implmented yet 
//	 * @param matchCase true to match the string cases; false to ignore string cases
//	 * @param sortByRows true to sort by rows(change columns orders); false to sort by columns(change row orders). 
//	 * @param sortMethod special sorting method
//	 * @param dataOption1 see numeric String as number or not for key1.
//	 * @param dataOption2 see numeric String as number or not for key2.
//	 * @param dataOption3 see numeric String as number or not for key3.
//	 */
//	public void sort(NRange rng1, boolean desc1, NRange rng2, int type, boolean desc2, NRange rng3, boolean desc3, int header, int orderCustom,
//			boolean matchCase, boolean sortByRows, int sortMethod, int dataOption1, int dataOption2, int dataOption3);

	/**
	 * Merge this range into a merged cell.
	 * @param across true to merge cells in each row; default to false.
	 */
	public void merge(boolean across);

	/**
	 * Un-merge merged cell in this range area to separated cells.
	 */
	public void unmerge();
	
//	/**
//	 * Adds/Remove border around this range.
//	 */
//	public void borderAround(BorderStyle lineStyle, String color);

	/**
	 * Adds/Remove border of all cell within this range per the specified border index.
	 * @param borderIndex one of {@link #BORDER_EDGE_BOTTOM},{@link #BORDER_EDGE_RIGHT},{@link #BORDER_EDGE_TOP},
	 * {@link #BORDER_EDGE_LEFT},{@link #BORDER_INSIDE_HORIZONTAL},{@link #BORDER_INSIDE_VERTICAL},{@link #BORDER_DIAGONAL_DOWN},
	 * {@link #BORDER_DIAGONAL_UP},{@link #BORDER_FULL},{@link #BORDER_OUTLINE},{@link #BORDER_INSIDE},{@link #BORDER_DIAGONAL}
 	 * @param lineStyle border line style, one of {@link BorderStyle} 
	 * @param color color in HTML format; i.e., #rrggbb.
	 */
	public void setBorders(ApplyBorderType borderIndex, NCellStyle.BorderType lineStyle, String color);

	/**
	 * Move this range to a new place as specified by nRow(negative value to move up; 
	 * positive value to move down) and nCol(negative value to move left; positive value to move right)
	 * @param nRow how many rows to move this range
	 * @param nCol how many columns to move this range
	 */
	public void move(int nRow, int nCol);
	
	/**
	 * Sets column width in unit of pixel
	 * @param widthPx 
	 */
	public void setColumnWidth(int widthPx);
	
	/**
	 * Sets row height in unit of pixel
	 * @param hightPx
	 */
	public void setRowHeight(int heightPx);

	/**
	 * Sets the width(in pixel) of column in this range, it effect to whole column. 
	 * @param widthPx width in pixel
	 * @param custom mark it as custom value
	 * @see #toColumnRange()
	 */
	public void setColumnWidth(int widthPx,boolean custom);
	/**
	 * Sets the height(in pixel) of row in this range, it effect to whole row.
	 * @param widthPx width in pixel
	 * @param custom mark it as custom value
	 * @see #toRowRange()
	 */
	public void setRowHeight(int heightPx,boolean custom);
	
	/**
	 * Returns associate {@link NSheet} of this range.
	 * @return associate {@link NSheet} of this range.
	 */
	public NSheet getSheet();
	
	/**
	 * Return the range that contains the cell specified in row, col (relative to this Range).
	 * @param row row index relative to this Range(note that it is 0-based)
	 * @param col column index relative to this Range(note that it is 0-based)
	 * @return the range that contains the cell specified in row, col (relative to this Range).
	 */
	public NRange getCells(int row, int col);
	
	/**
	 * Sets a Style object to this Range.
	 * @param style the style object
	 */
	public void setStyle(NCellStyle style);
	
	/**
	 * Perform an auto fill on the specified destination Range. Note the given destination Range
	 * must include this source Range.
	 * @param dstRange destination range to do the auto fill. Note the given destination Range must include this source Range
	 * @param fillType the fillType
	 */
	public void autoFill(NRange dstRange, AutoFillType fillType);
	
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
	 * To find a range of cells for applying auto filter according to this range.
	 * Usually, these two ranges are different.
	 * This method searches the filtering range through a specific rules. 
	 * @see org.zkoss.zss.api.Range#findAutoFilterRange()
	 * @return a range of cells for applying auto filter or null if can't find one from this Range. 
	 * @since 3.0.0
	 */
	// Refer to ZSS-246.
	NRange findAutoFilterRange();
	
	/**
	 * Filters a list specified by this Range and returns an AutoFilter object.
	 * @param field offset of the field on which you want to base the filter on (1-based; i.e. leftmost column in this range is field 1).
	 * @param filterOp, Use FILTEROP_AND and FILTEROP_OR with criteria1 and criterial2 to construct compound criteria.
	 * @param criteria1 "=" to find blank fields, "<>" to find non-blank fields. If null, means ALL. If filterOp == AutoFilter#FILTEROP_TOP10, 
	 * then this shall specifies the number of items (e.g. "10"). 
	 * @param criteria2 2nd criteria; used with criteria1 and filterOP to construct compound criteria.
	 * @param visibleDropDown true to show the autoFilter drop-down arrow for the filtered field; false to hide the autoFilter drop-down arrow; null
	 * to keep as is.
	 * @return the applied AutoFiltering
	 */
	public NAutoFilter enableAutoFilter(int field, FilterOp filterOp, Object criteria1, Object criteria2, Boolean visibleDropDown);
	
	/**
	 * Enable the auto filter and return it, get null if you disable it. 
	 * @return the autofilter if enable, or null if disable. 
	 */
	public NAutoFilter enableAutoFilter(boolean enable);

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
     * Sets the protection enabled as well as the password
     * @param password to set for protection. Pass <code>null</code> to remove protection
     */
    public void protectSheet(String password);
	
	/**
	 * Sets the hyperlink of this Range
	 * @param linkType the type of target to link. One of the {@link Hyperlink#LINK_URL}, 
	 * {@link Hyperlink#LINK_DOCUMENT}, {@link Hyperlink#LINK_EMAIL}, {@link Hyperlink#LINK_FILE}
	 * @param address the address
	 * @param display the text to display link
	 */
	public void setHyperlink(HyperlinkType linkType, String address, String display);
	
//	/**
//	 * Returns an {@link XAreas} which is a collection of each single selected area(also Range) of this multiple-selected Range. 
//	 * If this Range is a single selected Range, this method return the Areas which contains only this Range itself.
//	 * @return an {@link XAreas} which is a collection of each single selected area(also Range) of this multiple-selected Range.
//	 */
//	public XAreas getAreas();
	
	/**
	 * Returns a {@link NRange} that represent all columns of the 1st selected area of this Range. Note that only the 1st selected area is considered if this Range is a multiple-selected Range. 
	 * @return a {@link NRange} that represent all columns of this Range.
	 */
	public NRange getColumns();
	
	/**
	 * Returns a {@link NRange} that represent all rows of the 1st selected area of this Range. Note that only the 1st selected area is considered if this Range is a multiple-selected Range. 
	 * @return a {@link NRange} that represent all rows of this Range.
	 */
	public NRange getRows();

//	/**
//	 * Returns a {@link NRange} that represent all dependents of the left-top cell of the 1st selected area of this Range. 
//	 * Note that only the left-top cell of the 1st selected area is considered if this Range is a multiple-selected Range.
//	 * This could be multiple-selected Range if there are more than one dependent. 
//	 * @return a {@link NRange} that represent all dependents of the left-top cell of the 1st selected area of this Range.
//	 */
//	public NRange getDependents();
	
//	/**
//	 * Returns a {@link NRange} that represent all direct dependents of the left-top cell of the 1st selected area of this Range. 
//	 * Note that only the left-top cell of the 1st selected area is considered if this Range is a multiple-selected Range. 
//	 * This method could return multiple-selected Range if there are more than one dependent. 
//	 * @return a {@link NRange} that represent all direct dependents of the left-top cell of the 1st selected area of this Range.
//	 */
//	public NRange getDirectDependents();
	
//	/** 
//	 * Returns a {@link NRange} that represent all precedents of the left-top cell of the 1st selected area of this Range. 
//	 * Note that only the left-top cell of the 1st selected area is considered if this Range is a multiple-selected Range.
//	 * This method could return multiple-selected Range if there are more than one precedent. 
//	 * @return a {@link NRange} that represent all precedents of the left-top cell of the 1st selected area of this Range.
//	 */
//	public NRange getPrecedents();
//	
//	/**
//	 * Returns a {@link NRange} that represent all direct precedents of the left-top cell of the 1st selected area of this Range. 
//	 * Note that only the left-top cell of the 1st selected area is considered if this Range is a multiple-selected Range. 
//	 * This method could return multiple-selected Range if there are more than one precedent. 
//	 * @return a {@link NRange} that represent all direct precedents of the left-top cell of the 1st selected area of this Range.
//	 */
//	public NRange getDirectPrecedents();
	
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
	
//	/**
//	 * Returns the number of contained objects in this Range.
//	 * @return the number of contained objects in this Range.
//	 */
//	public long getCount();

	/**
	 * Set value into this Range.
	 * @param value the value
	 */
	public void setValue(Object value);
	
	/**
	 * Returns left top cell value of this Range.
	 * @return left top cell value of this Range
	 */
	public Object getValue();
	
	/**
	 * Returns a {@link NRange} that represents a range that offset from this Range. 
	 * @param rowOffset positive means downward; 0 means don't change row; negative means upward.
	 * @param colOffset positive means rightward; 0 means don't change column; negative means leftward.
	 * @return a {@link NRange} that represents a range that offset from this Range.
	 */
	public NRange getOffset(int rowOffset, int colOffset);
	
//	/**
//	 * Returns a {@link NRange} that bounds current Left-top cell of this Range with a combination of blank Rows and Columns.
//	 * @return a {@link NRange} that bounds current Left-top cell of this Range with a combination of blank Rows and Columns.
//	 */
//	public NRange getCurrentRegion();
//	
//	/**
//	 * Reapply current {@link AutoFilter}.
//	 */
//	public void applyFilter();
//	
//	/**
//	 * Clear all application of the current {@link AutoFilter}. 
//	 */
//	public void showAllData();

//	/**
//	 * Add a chart into the sheet of this Range 
//	 * @param anchor
//	 * @return the created chart 
//	 */
//	public Chart addChart(ClientAnchor anchor, ChartData data, ChartType type, ChartGrouping grouping, LegendPosition pos);
//
//	/**
//	 * Insert a picture into the sheet of this Range
//	 * @param anchor picture anchor
//	 * @param image image data
//	 * @param format image format
//	 * @see Workbook#PICTURE_TYPE_EMF
//     * @see Workbook#PICTURE_TYPE_WMF
//     * @see Workbook#PICTURE_TYPE_PICT
//     * @see Workbook#PICTURE_TYPE_JPEG
//     * @see Workbook#PICTURE_TYPE_PNG
//     * @see Workbook#PICTURE_TYPE_DIB
//     * @return the created picture
//	 */
//	public Picture addPicture(ClientAnchor anchor, byte[] image, int format);
//
//
//	/**
//	 * Delete an existing picture from the sheet of this Range.
//	 * @param picture the picture to be deleted
//	 */
//	public void deletePicture(Picture picture);
//	
//	/**
//	 * Update picture anchor.
//	 * @param picture the picture to change anchor
//	 * @param anchor the new anchor
//	 */
//	public void movePicture(Picture picture, ClientAnchor anchor);
//
//	/**
//	 * Update chart anchor.
//	 * @param chart the chart to change anchor
//	 * @param anchor the new anchor
//	 */
//	public void moveChart(Chart chart, ClientAnchor anchor);
//	
//	/**
//	 * Delete an existing chart from the sheet of this Range.
//	 * @param chart the chart to be deleted
//	 */
//	public void deleteChart(Chart chart);
//	
	/**
	 * Returns whether the plain text input by the end user is valid or not;
	 * note the validation only applies to the left-top cell of this Range.
	 * @param txt the string input by the end user.
	 * @return null if a valid input to the specified range; otherwise, the DataValidation
	 */
	public NDataValidation validate(String txt);
	
	/**
	 * Returns whether any cell is protected and locked in this Range.
	 * @return true if any cell is protected and locked in this Range.
	 */
	public boolean isAnyCellProtected();

//	/**
//	 * Move focus of the sheet of this Range(used for book collaboration).
//	 * @param token the token to identify the focus
//	 */
//	public void notifyMoveFriendFocus(Object token);
//	
//	/**
//	 * Delete focus of the sheet of this Range(used for book collaboration). 
//	 * @param token the token to identify the registration
//	 */
//	public void notifyDeleteFriendFocus(Object token);
	
	/**
	 * Delete sheet of this Range.
	 */
	public void deleteSheet();
	
	/**
	 * Create sheet of this book as specified in this Range.
	 * @param name the name of the new created sheet; null would use default 
	 * "SheetX" name where X is the next sheet number.
	 */
	public NSheet createSheet(String name);
	
	/**
     * Set(Rename) the name of the sheet as specified in this Range.
	 * @param name
	 */
	public void setSheetName(String name);
	
    /**
     * Sets the order of the sheet as specified in this Range.
     *
     * @param pos the position that we want to insert the sheet into (0 based)
     */
	public void setSheetOrder(int pos);
	
	/**
	 * Check if this range cover an entire rows (form 0, and last row to the max available row of a sheet) 
	 */
	public boolean isWholeRow();
	
	/**
	 * Check if this range cover an entire columns (form 0, and last row to the max available column of a sheet) 
	 */
	public boolean isWholeColumn();
	
	/**
	 * Check if this range cover an entire sheet 
	 */
	public boolean isWholeSheet();
	
	
	/**
	 * Notify this range has been changed.
	 */
	public void notifyChange();
	
	/**
	 * Set the freeze panel
	 * @param rowfreeze the number of row to freeze, 0 means no freeze
	 * @param columnfreeze the number of column to freeze, 0 means no freeze
	 */
	public void setFreezePanel(int rowfreeze, int columnfreeze);

	public String getCellFormatText();

	public boolean isSheetProtected();
	
}
