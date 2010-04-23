/* Range.java

	Purpose:
		
	Description:
		
	History:
		Mar 10, 2010 2:35:24 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model;

import java.util.List;
import java.util.Set;

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
	/**
	 * Returns formated text of this Range.
	 * @return formated text of this Range.
	 */
	public RichTextString getText();
	
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
}
