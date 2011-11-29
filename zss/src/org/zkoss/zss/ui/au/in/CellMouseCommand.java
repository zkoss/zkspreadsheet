/* CellMouseCommand.java

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
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zk.ui.event.MouseEvent;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.CellMouseEvent;
import org.zkoss.zss.ui.event.FilterMouseEvent;
import org.zkoss.zss.ui.impl.Utils;
import static org.zkoss.zss.ui.au.in.Commands.parseKeys;

/**
 * A Command (client to server) for handling user(client) start editing a cell
 * @author Dennis.Chen
 *
 */
public class CellMouseCommand implements Command {
	final static String Command = "onZSSCellMouse";
	//-- super --//
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, this);
		//final String[] data = request.getData();
		//TODO a little patch, "af" might shall be fired by a separate command?
		final Map data = (Map) request.getData();
		String type = (String) data.get("type");//command type
		if (data == null 
			|| (!"af".equals(type) && !"dv".equals(type) && data.size() != 9) 
			|| (("af".equals(type) || "dv".equals(type)) && data.size() != 10))
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), this});
		
		int shx = (Integer) data.get("shx");//x offset against spreadsheet
		int shy = (Integer) data.get("shy");
		int key = parseKeys((String) data.get("key"));
		String sheetId = (String) data.get("sheetId");
		int row = (Integer) data.get("row");
		int col = (Integer) data.get("col");
		int mx = (Integer) data.get("mx");//x offset against body
		int my = (Integer) data.get("my");
		
		Worksheet sheet = ((Spreadsheet) comp).getSelectedSheet();
		if (!Utils.getSheetUuid(sheet).equals(sheetId))
			return;
		
		if ("lc".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_CELL_CLICK;
		} else if ("rc".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_CELL_RIGHT_CLICK;
		} else if ("dbc".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_CELL_DOUBLE_CLICK;
		} else if ("af".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_FILTER;
		} else if ("dv".equals(type)) {
			type = org.zkoss.zss.ui.event.Events.ON_VALIDATE_DROP;
		} else {
			throw new UiException("unknow type : " + type);
		}

		if (org.zkoss.zss.ui.event.Events.ON_FILTER.equals(type)) {
			int field = (Integer) data.get("field");
			Events.postEvent(new FilterMouseEvent(type, comp, shx, shy, key, sheet, row, col, mx, my, field));
		} else if (org.zkoss.zss.ui.event.Events.ON_VALIDATE_DROP.equals(type)) {
			int dvindex = (Integer) data.get("field");
			Events.postEvent(new FilterMouseEvent(type, comp, shx, shy, key, sheet, row, col, mx, my, dvindex));
		} else {
			Events.postEvent(new CellMouseEvent(type, comp, shx, shy, key, sheet, row, col, mx, my));
		}
	}
	
	public String getCommand() {
		return Command;
	}
}