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

import org.zkoss.json.JSONObject;
import org.zkoss.lang.Objects;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.SheetCtrl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;

/**
 * @author sam
 *
 */
public class FetchActiveRangeCommand extends AbstractCommand implements Command {

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
		
		XSheet sheet = spreadsheet.getSelectedXSheet();
		
		if (sheetId.equals(((SheetCtrl)sheet).getUuid())) {
			final SpreadsheetCtrl spreadsheetCtrl = ((SpreadsheetCtrl) spreadsheet.getExtraCtrl());
			
			JSONObject mainBlock = spreadsheetCtrl.getRangeAttrs(sheet, SpreadsheetCtrl.Header.BOTH, SpreadsheetCtrl.CellAttribute.ALL, left, top, right, bottom);
			if (spreadsheet.getColumnfreeze() > -1) {
				mainBlock.put("leftFrozen", 
						spreadsheetCtrl.getRangeAttrs(sheet, SpreadsheetCtrl.Header.BOTH, SpreadsheetCtrl.CellAttribute.ALL, 0, top, spreadsheet.getColumnfreeze(), bottom));
			}
			if (spreadsheet.getRowfreeze() > -1) {
				mainBlock.put("topFrozen", 
						spreadsheetCtrl.getRangeAttrs(sheet, SpreadsheetCtrl.Header.BOTH, SpreadsheetCtrl.CellAttribute.ALL, left, 0, right, spreadsheet.getRowfreeze()));
			}
			spreadsheet.smartUpdate("activeRangeUpdate", mainBlock);
		}
	}
}