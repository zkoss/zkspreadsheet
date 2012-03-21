/* InnerEvts.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 1, 2012 10:25:55 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui;

import java.util.HashMap;
import java.util.Map;

import org.zkoss.zss.ui.au.in.ActionCommand;
import org.zkoss.zss.ui.au.in.BlockSyncCommand;
import org.zkoss.zss.ui.au.in.CellFetchCommand;
import org.zkoss.zss.ui.au.in.CellFocusedCommand;
import org.zkoss.zss.ui.au.in.CellMouseCommand;
import org.zkoss.zss.ui.au.in.CellSelectionCommand;
import org.zkoss.zss.ui.au.in.Command;
import org.zkoss.zss.ui.au.in.CtrlKeyCommand;
import org.zkoss.zss.ui.au.in.EditboxEditingCommand;
import org.zkoss.zss.ui.au.in.FetchActiveRangeCommand;
import org.zkoss.zss.ui.au.in.FilterCommand;
import org.zkoss.zss.ui.au.in.HeaderCommand;
import org.zkoss.zss.ui.au.in.HeaderMouseCommand;
import org.zkoss.zss.ui.au.in.MoveWidgetCommand;
import org.zkoss.zss.ui.au.in.SelectSheetCommand;
import org.zkoss.zss.ui.au.in.SelectionChangeCommand;
import org.zkoss.zss.ui.au.in.StartEditingCommand;
import org.zkoss.zss.ui.au.in.StopEditingCommand;
import org.zkoss.zss.ui.au.in.WidgetCtrlKeyCommand;

/**
 * @author sam / Ian
 *
 */
/*package*/ final class InnerEvts {
	private InnerEvts() {}
	
	static final String ON_ZSS_ACTION = "onZSSAction";
	static final String ON_ZSS_CELL_FETCH = "onZSSCellFetch";
	static final String ON_ZSS_CELL_FOCUSED = "onZSSCellFocused";
	static final String ON_ZSS_CELL_MOUSE = "onZSSCellMouse";
	static final String ON_ZSS_CELL_SELECTION = org.zkoss.zss.ui.event.Events.ON_CELL_SELECTION;
	static final String ON_ZSS_CELL_SELECTION_CHANGE = org.zkoss.zss.ui.event.Events.ON_SELECTION_CHANGE;
	static final String ON_ZSS_EDITBOX_EDITING = org.zkoss.zss.ui.event.Events.ON_EDITBOX_EDITING;
	static final String ON_ZSS_FETCH_ACTIVE_RANGE = "onZSSFetchActiveRange";
	static final String ON_ZSS_FILTER = "onZSSFilter";
	static final String ON_ZSS_HEADER_MODIF = "onZSSHeaderModif";
	static final String ON_ZSS_HEADER_MOUSE = "onZSSHeaderMouse";
	static final String ON_ZSS_MOVE_WIDGET = "onZSSMoveWidget";
	static final String ON_ZSS_WIDGET_CTRL_KEY = "onZSSWidgetCtrlKey";
	static final String ON_ZSS_SELECT_SHEET = "onZSSSelectSheet";
	static final String ON_ZSS_START_EDITING = org.zkoss.zss.ui.event.Events.ON_START_EDITING;
	static final String ON_ZSS_STOP_EDITING = org.zkoss.zss.ui.event.Events.ON_STOP_EDITING;
	static final String ON_ZSS_SYNC_BLOCK = "onZSSSyncBlock";
	static final String ON_ZSS_CTRL_KEY = org.zkoss.zk.ui.event.Events.ON_CTRL_KEY;
	
	static final Map<String, Command> CMDS;
	static{
		CMDS = new HashMap<String, Command>();
		CMDS.put(ON_ZSS_ACTION, new ActionCommand());
		CMDS.put(ON_ZSS_CELL_FETCH, new CellFetchCommand());
		CMDS.put(ON_ZSS_CELL_FOCUSED, new CellFocusedCommand());
		CMDS.put(ON_ZSS_CELL_MOUSE, new CellMouseCommand());
		CMDS.put(ON_ZSS_CELL_SELECTION, new CellSelectionCommand());
		CMDS.put(ON_ZSS_CELL_SELECTION_CHANGE, new SelectionChangeCommand());
		CMDS.put(ON_ZSS_EDITBOX_EDITING, new EditboxEditingCommand());
		CMDS.put(ON_ZSS_FETCH_ACTIVE_RANGE, new FetchActiveRangeCommand());
		CMDS.put(ON_ZSS_FILTER, new FilterCommand());
		CMDS.put(ON_ZSS_HEADER_MODIF, new HeaderCommand());
		CMDS.put(ON_ZSS_HEADER_MOUSE, new HeaderMouseCommand());
		CMDS.put(ON_ZSS_MOVE_WIDGET, new MoveWidgetCommand());
		CMDS.put(ON_ZSS_WIDGET_CTRL_KEY, new WidgetCtrlKeyCommand());
		CMDS.put(ON_ZSS_SELECT_SHEET, new SelectSheetCommand());
		CMDS.put(ON_ZSS_START_EDITING, new StartEditingCommand());
		CMDS.put(ON_ZSS_STOP_EDITING, new StopEditingCommand());
		CMDS.put(ON_ZSS_SYNC_BLOCK, new BlockSyncCommand());
		CMDS.put(ON_ZSS_CTRL_KEY, new CtrlKeyCommand());
	}
	/**
	 * 
	 * @param name
	 * @return
	 */
	static Command getCommand(String name){
		return CMDS.get(name);
	}
}
