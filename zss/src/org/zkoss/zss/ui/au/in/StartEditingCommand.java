/* StartEditingCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 18, 2007 12:10:40 PM     2007, Created by Dennis.Chen
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
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.RichTextString;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.StartEditingEvent;
import org.zkoss.zss.ui.impl.XUtils;


/**
 * A Command (client to server) for handling user(client) start editing a cell
 * @author Dennis.Chen
 *
 */
public class StartEditingCommand implements Command {

	public void process (AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, this);
		final Map data = (Map)request.getData();
		if (data == null || data.size() != 6)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), this});

		String token = (String) data.get("token");
		String sheetId = (String) data.get("sheetId");
		int row = (Integer) data.get("row");
		int col = (Integer) data.get("col");
		String clienttxt = (String) data.get("clienttxt");
		String type = (String) data.get("type");

		XSheet sheet = ((Spreadsheet) comp).getSelectedXSheet();
		if (!XUtils.getSheetUuid(sheet).equals(sheetId))
			return;
		
		Cell cell = XUtils.getCell(sheet, row, col);
		// You can call getEditText(), setEditText(), getText().

		String editText;
		if(cell==null){
			editText = "";
		}else{
			RichTextString rts = BookHelper.getRichEditText(cell);
			if(rts==null){
				editText = "";
			}else{
				editText = rts.getString();
			}
		}
				
		
		StartEditingEvent event = new StartEditingEvent(
				org.zkoss.zss.ui.event.Events.ON_START_EDITING, comp, sheet,
				row, col, editText, clienttxt);
		Events.postEvent(event);
		Events.postEvent(new Event("onStartEditingImpl", comp, new Object[] {token, event, type}));
	}
}