/* CellRichTextAction.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2014/08/20, Created by henrichen
}}IS_NOTE

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl.undo;

import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
/**
 * 
 * @author henrichen
 * @since 3.5.1
 */
public class CellRichTextAction extends AbstractEditTextAction {
	
	private final String _richText;
	
	public CellRichTextAction(String label,Sheet sheet,int row, int column, int lastRow,int lastColumn,String editText){
		super(label,sheet,row,column,lastRow,lastColumn);
		this._richText = editText;
	}

	protected void applyAction(){
		boolean protect = isSheetProtected();
		if(_richText!=null && !protect){
			Range r = Ranges.range(_sheet,_row,_column,_lastRow,_lastColumn);
			r.setCellRichText(_richText);
		}
	}
}
