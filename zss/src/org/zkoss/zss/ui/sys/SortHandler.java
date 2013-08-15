/* SortHandler.java

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
package org.zkoss.zss.ui.sys;

import org.zkoss.util.resource.Labels;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.ui.impl.ua.AbstractProtectedHandler;
import org.zkoss.zss.ui.impl.undo.SortCellAction;

/**
 * @author dennis
 *
 */
public class SortHandler extends AbstractProtectedHandler {

	boolean _desc;
	public SortHandler(boolean desc) {
		this._desc = desc;
	}


	/* (non-Javadoc)
	 * @see org.zkoss.zss.ui.sys.ua.impl.AbstractHandler#processAction(org.zkoss.zss.ui.UserActionContext)
	 */
	@Override
	protected boolean processAction(UserActionContext ctx) {
		Sheet sheet = ctx.getSheet();
		AreaRef selection = ctx.getSelection();
		UndoableActionManager uam = ctx.getSpreadsheet().getUndoableActionManager();
		uam.doAction(new SortCellAction(Labels.getLabel("zss.undo.sortAsc"),sheet, selection.getRow(), selection.getColumn(),
			selection.getLastRow(), selection.getLastColumn(),_desc));
		ctx.clearClipboard();
		return true;
	}

}
