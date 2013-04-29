/* SelectionChangeCommand.java
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
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.Events;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.SelectionChangeEvent;
import org.zkoss.zss.ui.event.CellSelectionEvent;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zss.ui.sys.SpreadsheetInCtrl;
/**
 * A Command (client to server) for handling cell selection
 * @author Dennis.Chen
 *
 */
public class SelectionChangeCommand implements Command {

	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, this);
		final Map data = (Map) request.getData();
		if (data == null || data.size() != 10)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), this});
		
		String sheetId= (String) data.get("sheetId");
		
		XSheet sheet = ((Spreadsheet) comp).getSelectedXSheet();
		if (!Utils.getSheetUuid(sheet).equals(sheetId))
			return;
		
		int action = (Integer) data.get("action");
		int left = (Integer) data.get("left");
		int top = (Integer) data.get("top");
		int right = (Integer) data.get("right");
		int bottom = (Integer) data.get("bottom");
		int orgileft = (Integer) data.get("orgileft");
		int orgitop = (Integer) data.get("orgitop");
		int orgiright = (Integer) data.get("orgiright");
		int orgibottom = (Integer) data.get("orgibottom");
		
		final SelectionChangeEvent evt = new SelectionChangeEvent(
				org.zkoss.zss.ui.event.Events.ON_SELECTION_CHANGE, comp, sheet,
				action, left, top, right, bottom, orgileft, orgitop, orgiright,
				orgibottom);
		
		if (!isProtect(top, left, bottom, right, sheet)) {
			final int xaction = evt.getAction();
			if (xaction == SelectionChangeEvent.MOVE) {
				final int nRow = top - orgitop;
				final int nCol = left - orgileft;
				
				switch(evt.getSelectionType()) {
				case CellSelectionEvent.SELECT_ROW:
					Utils.moveRows(sheet, orgitop, orgibottom, nRow);
					break;
				case CellSelectionEvent.SELECT_COLUMN:
					Utils.moveColumns(sheet, orgileft, orgiright, nCol);
					break;
				case CellSelectionEvent.SELECT_CELLS:
					Utils.moveCells(sheet, orgitop, orgileft, orgibottom, orgiright, nRow, nCol);
					break;
				}
			} else if (xaction == SelectionChangeEvent.MODIFY) {
				switch(evt.getSelectionType()) {
				case CellSelectionEvent.SELECT_ROW:
					Utils.fillRows(sheet, orgitop, orgibottom, top, bottom);
					break;
				case CellSelectionEvent.SELECT_COLUMN:
					Utils.fillColumns(sheet, orgileft, orgiright, left, right);
					break;
				case CellSelectionEvent.SELECT_CELLS:
					Utils.fillCells(sheet, orgitop, orgileft, orgibottom, orgiright, top, left, bottom, right);
					break;
				}
			}	
		}
		SpreadsheetInCtrl ctrl = ((SpreadsheetInCtrl)((Spreadsheet)comp).getExtraCtrl());
		ctrl.setSelectionRect(left, top, right, bottom);	
		
		Events.postEvent(evt);
	}
	
	private boolean isProtect(int tRow, int lCol, int bRow, int rCol, XSheet sheet) {
		boolean shtProtect = sheet.getProtect();
		if (!shtProtect)
			return false;
		
		for (int r = tRow; r <= bRow; r++) {
			Row row = sheet.getRow(r);
			if (row != null) {
				for (int c = lCol; c <= rCol; c++) {
					Cell cell = row.getCell(c);
					if (shtProtect && cell != null && cell.getCellStyle().getLocked()) {
						return true;
					} else if (shtProtect && cell == null) {
						return true;
					}
				}	
			} else if (shtProtect && row == null) {
				return true;
			}
		}
		return false;
	}
}