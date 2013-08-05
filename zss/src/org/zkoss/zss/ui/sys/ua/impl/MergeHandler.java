/* ToggleMergeCenterHandler.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/8/5 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.sys.ua.impl;

import org.zkoss.util.resource.Labels;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.undo.MergeCellAction;
import org.zkoss.zss.undo.UndoableActionManager;

/**
 * @author dennis
 *
 */
public class MergeHandler extends AbstractProtectedHandler {

	boolean _across;
	
	public MergeHandler(boolean across) {
		this._across = across;
	}




	/* (non-Javadoc)
	 * @see org.zkoss.zss.ui.sys.ua.impl.AbstractHandler#processAction(org.zkoss.zss.ui.UserActionContext)
	 */
	@Override
	protected boolean processAction(UserActionContext ctx) {
		Sheet sheet = ctx.getSheet();
		Rect selection = ctx.getSelection();
		Range range = Ranges.range(sheet, selection);
		UndoableActionManager uam = ctx.getSpreadsheet().getUndoableActionManager();
		uam.doAction(new MergeCellAction(Labels.getLabel("zss.undo.merge"),sheet, selection.getRow(), selection.getColumn(), 
			selection.getLastRow(), selection.getLastColumn(),_across));
		
		ctx.clearClipboard();
		return true;
	}

}
