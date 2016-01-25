/* PasteRangeImpl.java

	Purpose:
		
	Description:
		
	History:
		Jan 22, 2016 5:01:51 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import org.zkoss.zss.model.SSheet;

/**
 * To avoid paste too many rows; ZSS copy column only copy to 
 * sheetMaxVisibleRows.
 * 
 * @author henri
 *
 */
public class PasteRangeImpl extends RangeImpl {
	private boolean _wholeRow;
	private boolean _wholeColumn;
	
	public PasteRangeImpl(SSheet sheet, int tRow, int lCol, int bRow, int rCol, boolean wholeRow, boolean wholeColumn) {
		super(sheet, tRow, lCol, bRow, rCol);
		_wholeRow = wholeRow;
		_wholeColumn = wholeColumn;
	}

	@Override
	public boolean isWholeSheet(){
		return (_wholeRow && _wholeColumn) || super.isWholeSheet();
	}

	@Override
	public boolean isWholeRow() {
		return _wholeRow  || super.isWholeRow();
	}
	
	@Override
	public boolean isWholeColumn() {
		return _wholeColumn || super.isWholeColumn(); 
	}
}
