/* FormatText.java

	Purpose:
		
	Description:
		
	History:
		Jun 15, 2010 11:26:06 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.model.impl;

import org.apache.poi.ss.format.CellFormatResult;
import org.apache.poi.ss.usermodel.RichTextString;
import org.zkoss.zss.model.FormatText;

/**
 * A cell formatted result holder class. 
 * @author henrichen
 *
 */
public class FormatTextImpl implements FormatText {
	private final CellFormatResult _result;
	private final RichTextString _rstr;
	
	public FormatTextImpl(CellFormatResult result) {
		_result = result;
		_rstr = null;
	}
	
	public FormatTextImpl(RichTextString rstr) {
		_rstr = rstr;
		_result = null;
	}
	
	public boolean isRichTextString() {
		return _rstr != null;
	}
	
	public boolean isCellFormatResult() {
		return _result != null;
	}
	
	public CellFormatResult getCellFormatResult() {
		return _result;
	}
	
	public RichTextString getRichTextString() {
		return _rstr;
	}

}
