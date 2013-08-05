package org.zkoss.zss.ui.sys.ua.impl;

import org.zkoss.util.resource.Labels;
import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.undo.HideHeaderAction;
import org.zkoss.zss.undo.HideHeaderAction.Type;
import org.zkoss.zss.undo.UndoableActionManager;

public class HideHeaderHandler extends AbstractProtectedHandler {
	final HideHeaderAction.Type _type;
	final boolean _hide;
	
	public HideHeaderHandler(Type type, boolean hide) {
		this._type = type;
		this._hide = hide;
	}



	@Override
	protected boolean processAction(UserActionContext ctx) {
		Sheet sheet = ctx.getSheet();
		Rect selection = ctx.getSelection();
		Range range = Ranges.range(sheet, selection);
		if(range.isProtected()){
			showProtectMessage();
			return true;
		}
		
		String label = null;
		switch(_type){
		case COLUMN:
			label = _hide?Labels.getLabel("zss.undo.hideColumn"):Labels.getLabel("zss.undo.unhideColumn");
			break;
		case ROW:
			label = _hide?Labels.getLabel("zss.undo.hideRow"):Labels.getLabel("zss.undo.unhideRow");
			break;
		}
		
		UndoableActionManager uam = ctx.getSpreadsheet().getUndoableActionManager();
		uam.doAction(new HideHeaderAction(label,sheet, selection.getRow(), selection.getColumn(), 
			selection.getLastRow(), selection.getLastColumn(), 
			_type,_hide));
		
		return true;
	}

}
