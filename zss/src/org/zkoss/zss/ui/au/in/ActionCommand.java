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
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.SpreadsheetLabel;
import org.zkoss.zss.ui.impl.Utils;

/**
 * @author sam
 *
 */
public class ActionCommand implements Command {

	@Override
	public String getCommand() {
		return "onZSSAction";
	}

	@Override
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, ActionCommand.class);
		
		final Map data = (Map) request.getData();
		if (data == null || data.size() != 4)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), ActionCommand.class });
		
		Spreadsheet spreadsheet = ((Spreadsheet) comp);
		
		String sheetId = (String) data.get("sheetId");
		Worksheet sheet = Utils.getSheetByUuid(spreadsheet.getBook(), sheetId);
		if (sheet != null) {
			String tag = (String) data.get("tag");
			String act = (String) data.get("act");
			String value = (String) data.get("value");
			if ("sheet".equals(tag)) {
				processSheet(act, value, sheet, spreadsheet);
			}
		}
	}
	
	
	private void processSheet(String action, String value, Worksheet sheet, Spreadsheet spreadsheet) {
		Book book = spreadsheet.getBook();
		if ("add".equals(action)) {
			String prefix = Labels.getLabel(SpreadsheetLabel.Sheet.SHEET.getLabelKey());
			if (Strings.isEmpty(prefix))
				prefix = "Sheet";
			int numSheet = book.getNumberOfSheets();
			Ranges.range(sheet).createSheet(prefix + " " + (numSheet + 1));
		} else if ("delete".equals(action)) {
			int numSheet = book.getNumberOfSheets();
			if (numSheet > 1) {
				Worksheet sel = null;
				int index = book.getSheetIndex(sheet);
				if (index == numSheet - 1) {//delete last sheet, move select sheet left
					sel = book.getWorksheetAt(index - 1);
				} else { //move sheet right
					sel = book.getWorksheetAt(index + 1);
				}
				Ranges.range(sheet).deleteSheet();
				spreadsheet.setSelectedSheet(sel.getSheetName());
			}
		} else if ("rename".equals(action)) {
			Ranges.range(sheet).setSheetName(value);
		} else if ("protect".equals(action)) {
			boolean protect = sheet.getProtect();
			Ranges.range(sheet).protectSheet(protect ? null : "");// toggle sheet protect
		} else if ("moveLeft".equals(action)) {
			int index = book.getSheetIndex(sheet);
			if (index > 0) {
				Ranges.range(sheet).setSheetOrder(index - 1);
			}
		} else if ("moveRight".equals(action)) {
			int index = book.getSheetIndex(sheet);
			if (index < book.getNumberOfSheets() - 1) {
				Ranges.range(sheet).setSheetOrder(index + 1);
			}
		}
	}
}
