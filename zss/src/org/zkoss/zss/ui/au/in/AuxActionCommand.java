/* AuxActionCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 13, 2012 12:37:46 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.au.in;

import java.util.Map;

import org.zkoss.lang.Objects;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.DefaultUserAction;
import org.zkoss.zss.ui.event.AuxActionEvent;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.impl.XUtils;

/**
 * @author sam
 * @author dennis
 */
public class AuxActionCommand implements Command {

	@Override
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, AuxActionCommand.class);
		
		final Map data = (Map) request.getData();
		if (data == null || data.size() < 2)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), AuxActionCommand.class });
		
		Spreadsheet spreadsheet = ((Spreadsheet) comp);
		String tag = (String) data.get("tag");
		String action = (String) data.get("act");
		Rect selection = getSelectionIfAny(data);
		Sheet sheet = null;
		
		if(selection==null){
			selection = spreadsheet.getSelection();
		}
		
		//old code logic refer to ActionCommand in 2.6.0
		if ("sheet".equals(tag) && spreadsheet.getXBook() != null) {
			String sheetId = (String) data.get("sheetId");
			XSheet xsheet = XUtils.getSheetByUuid(spreadsheet.getXBook(), sheetId);
			
			if(xsheet==null){
				throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), AuxActionCommand.class });
			}

			// get back sheet by xsheet's name
			sheet = spreadsheet.getBook().getSheet(xsheet.getSheetName());

			// client's act doesn't follow the Action, so I have to remap it.
			// TODO make client use correct key directly?
			if ("add".equals(action)) {
				action = DefaultUserAction.ADD_SHEET.toString();
			} else if ("delete".equals(action)) {
				action = DefaultUserAction.DELETE_SHEET.toString();
			} else if ("rename".equals(action)) {
				action = DefaultUserAction.RENAME_SHEET.toString();
			} else if ("protect".equals(action)) {
				action = DefaultUserAction.PROTECT_SHEET.toString();
			} else if ("moveLeft".equals(action)) {
				action = DefaultUserAction.MOVE_SHEET_LEFT.toString();
			} else if ("moveRight".equals(action)) {
				action = DefaultUserAction.MOVE_SHEET_RIGHT.toString();
			} else {
				throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA,
						new Object[] { Objects.toString(data),
								AuxActionCommand.class });
			}
		}else if ("toolbar".equals(tag)) {
			sheet = spreadsheet.getSelectedSheet();
		}else{
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), AuxActionCommand.class });
		}
		
		AuxActionEvent evt = new AuxActionEvent(Events.ON_AUX_ACTION, spreadsheet, sheet, action, selection,data);
		
		org.zkoss.zk.ui.event.Events.postEvent(evt);
	}
	
	private Rect getSelectionIfAny(Map data) {
		if(data.containsKey("tRow") && data.containsKey("tRow") && data.containsKey("tRow") && data.containsKey("tRow")){
			int tRow = (Integer) data.get("tRow");
			int bRow = (Integer) data.get("bRow");
			int lCol = (Integer) data.get("lCol");
			int rCol = (Integer) data.get("rCol");
			Integer action = (Integer) data.get("action");
//			Rect r = action != null ? 
//					new Rect(action, lCol, tRow, rCol, bRow) : new Rect(lCol, tRow, rCol, bRow);
			Rect r = new Rect(lCol, tRow, rCol, bRow);
			return r;
		}else{
			return null;
		}
	}
}
