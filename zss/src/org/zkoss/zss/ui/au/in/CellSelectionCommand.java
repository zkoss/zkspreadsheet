/* CellSelectionCommand.java
 * 
{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		January 10, 2008 03:10:40 PM , Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under Lesser GPL Version 2.1 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;


import java.util.Map;

import org.zkoss.lang.Objects;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Events;
//import org.zkoss.zss.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.sys.SpreadsheetInCtrl;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * A Command (client to server) for handling cell selection
 * @author Dennis.Chen
 *
 */
public class CellSelectionCommand implements Command {
	final static String Command = org.zkoss.zss.ui.event.Events.ON_CELL_SELECTION;

	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, CellSelectionCommand.class.getCanonicalName());

		final Map data = (Map) request.getData();
		if (data == null || data.size() != 6)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), this});
			
		String sheetId= (String) data.get("sheetId");
		
		Sheet sheet = ((Spreadsheet)comp).getSelectedSheet();
		if(!Utils.getSheetId(sheet).equals(sheetId))
			return;
		
		int action = (Integer) data.get("action");
		int left = (Integer) data.get("left");
		int top = (Integer) data.get("top");
		int right = (Integer) data.get("right");
		int bottom = (Integer) data.get("bottom");
		
		SpreadsheetInCtrl ctrl = ((SpreadsheetInCtrl)((Spreadsheet)comp).getExtraCtrl());
		ctrl.setSelectionRect(left, top, right, bottom);	
		
		Events.postEvent(new CellSelectionEvent(org.zkoss.zss.ui.event.Events.ON_CELL_SELECTION, comp, sheet,action,left,top,right,bottom));
	}

	public String getCommand() {
		return Command;
	}
}