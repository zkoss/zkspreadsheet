/* MenuitemCommand.java

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
import org.zkoss.lang.Strings;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.UserAction;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.XUtils;

/**
 * @author sam
 */
public class ActionCommand implements Command {

	@Override
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, ActionCommand.class);
		
		final Map data = (Map) request.getData();
		if (data == null || data.size() < 2)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), ActionCommand.class });
		
		Spreadsheet spreadsheet = ((Spreadsheet) comp);
		String tag = (String) data.get("tag");
		String act = (String) data.get("act");
		if ("toolbar".equals(tag)) {
			//toolbar's are always target on selectedSheet
			spreadsheet.getUserActionHandler().handleAction(spreadsheet, spreadsheet.getSelectedSheet(), act, getSelectionIfAny(data), data);
		} else if ("sheet".equals(tag) && spreadsheet.getXBook() != null) {
			String sheetId = (String) data.get("sheetId");
			XSheet xsheet = XUtils.getSheetByUuid(spreadsheet.getXBook(), sheetId);
			if (xsheet != null) {
				//get back sheet by xsheet's name
				Sheet sheet = spreadsheet.getBook().getSheet(xsheet.getSheetName());
				
				//client's act doesn't follow the Action, so I have to remap it.
				if ("add".equals(act)) {
					act = UserAction.ADD_SHEET.toString();
				} else if ("delete".equals(act)) {
					act = UserAction.DELETE_SHEET.toString();
				} else if ("rename".equals(act)) {
					act = UserAction.RENAME_SHEET.toString();
				} else if ("protect".equals(act)) {
					act = UserAction.PROTECT_SHEET.toString();
				} else if ("moveLeft".equals(act)) {
					act = UserAction.MOVE_SHEET_LEFT.toString();
				} else if ("moveRight".equals(act)) {
					act = UserAction.MOVE_SHEET_RIGHT.toString();
				}
				spreadsheet.getUserActionHandler().handleAction(spreadsheet, sheet, act, getSelectionIfAny(data), data);
			}
		}
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
	
//	private void processSheet(String action, Map data, XSheet sheet, Spreadsheet spreadsheet) {
//		XBook book = spreadsheet.getXBook();
//		if ("add".equals(action)) {
//			String prefix = Labels.getLabel(Action.SHEET.getLabelKey());
//			if (Strings.isEmpty(prefix))
//				prefix = "Sheet";
//			int numSheet = book.getNumberOfSheets();
//			XRanges.range(sheet).createSheet(prefix + " " + (numSheet + 1));
//		} else if ("delete".equals(action)) {
//			int numSheet = book.getNumberOfSheets();
//			if (numSheet > 1) {
//				XSheet sel = null;
//				int index = book.getSheetIndex(sheet);
//				if (index == numSheet - 1) {//delete last sheet, move select sheet left
//					sel = book.getWorksheetAt(index - 1);
//				} else { //move sheet right
//					sel = book.getWorksheetAt(index + 1);
//				}
//				XRanges.range(sheet).deleteSheet();
//				spreadsheet.setSelectedSheet(sel.getSheetName());
//			}
//		} else if ("rename".equals(action)) {
//			String name = (String) data.get("name");
//			XRanges.range(sheet).setSheetName(name);
//		} else if ("protect".equals(action)) {
//			boolean protect = sheet.getProtect();
//			XRanges.range(sheet).protectSheet(protect ? null : "");// toggle sheet protect
//		} else if ("moveLeft".equals(action)) {
//			int index = book.getSheetIndex(sheet);
//			if (index > 0) {
//				XRanges.range(sheet).setSheetOrder(index - 1);
//			}
//		} else if ("moveRight".equals(action)) {
//			int index = book.getSheetIndex(sheet);
//			if (index < book.getNumberOfSheets() - 1) {
//				XRanges.range(sheet).setSheetOrder(index + 1);
//			}
//		}
//	}
}
