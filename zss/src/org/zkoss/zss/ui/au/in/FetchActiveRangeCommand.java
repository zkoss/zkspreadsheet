/* DeferLoadingCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 9, 2012 10:29:15 AM , Created by sam
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
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.SheetCtrl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;

/**
 * @author sam
 *
 */
public class FetchActiveRangeCommand implements Command {

	@Override
	public String getCommand() {
		return "onZSSFetchActiveRange";
	}

	@Override
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED,
					FetchActiveRangeCommand.class);
		
		final Map data = (Map) request.getData();
		if (data == null || data.size() != 5)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), SelectSheetCommand.class });
		
		Spreadsheet spreadsheet = ((Spreadsheet) comp);
		String sheetId = (String) data.get("sheetId");
		int left = (Integer) data.get("left");
		int right = (Integer) data.get("right");
		int top = (Integer) data.get("top");
		int bottom = (Integer) data.get("bottom");
		
		Worksheet sheet = spreadsheet.getSelectedSheet();
		
		if (sheetId.equals(((SheetCtrl)sheet).getUuid())) {
			final SpreadsheetCtrl spreadsheetCtrl = ((SpreadsheetCtrl) spreadsheet.getExtraCtrl());
			
			spreadsheet.smartUpdate("activeRangeUpdate",
				spreadsheetCtrl.getRangeAttrs(sheet, SpreadsheetCtrl.Header.BOTH, SpreadsheetCtrl.CellAttribute.ALL, left, top, right, bottom));
		}
	}
}
