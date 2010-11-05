/* FormatText.java

	Purpose:
		
	Description:
		
	History:
		Jun 15, 2010 11:30:48 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model;

import org.zkoss.poi.ss.format.CellFormatResult;
import org.zkoss.poi.ss.usermodel.RichTextString;

/**
 * Holding class for possible cell formatted text results
 * @author henrichen
 *
 */
public interface FormatText {
	/**
	 * Returns whether the formatted text a RichTextString.
	 * @return whether a RichTextString.
	 */
	public boolean isRichTextString();
	
	/**
	 * Returns whether the formatted text a CellFormatResult.
	 * @return whether a CellFormatResult.
	 */
	public boolean isCellFormatResult();
	
	/**
	 * Returns the CellFormatResult.
	 * @return the CellFormatResult.
	 */
	public CellFormatResult getCellFormatResult();
	
	/**
	 * Returns the RichTextString.
	 * @return the RichTextString.
	 */
	public RichTextString getRichTextString();
}
