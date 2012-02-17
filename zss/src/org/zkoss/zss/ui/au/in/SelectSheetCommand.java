/* SelectSheetCommand.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 2, 2012 10:36:09 AM , Created by sam
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
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.SheetCtrl;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.event.Events;

/**
 * @author sam
 *
 */
public class SelectSheetCommand implements Command {

	@Override
	public String getCommand() {
		return "onZSSSelectSheet";
	}
	
	@Override
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED,
					SelectSheetCommand.class);
		
		final Map data = (Map) request.getData();
		if (data == null || data.size() != 12)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), SelectSheetCommand.class });
		
		Spreadsheet spreadsheet = ((Spreadsheet) comp);
		String sheetId = (String) data.get("sheetId");
		boolean cacheInClient = (Boolean) data.get("cache");
		int row = (Integer)data.get("row");
		int col = (Integer)data.get("col");
		
		//selection
		int top = (Integer)data.get("top");
		int right = (Integer)data.get("right");
		int bottom = (Integer)data.get("bottom");
		int left = (Integer)data.get("left");
		
		//highlight
		int highlightLeft = (Integer)data.get("hleft");
		int highlightTop = (Integer)data.get("htop");
		int highlightRight = (Integer)data.get("hright");
		int highlightBottom = (Integer)data.get("hbottom");
		
		Book book = spreadsheet.getBook();
		int len = book.getNumberOfSheets();
		for (int i = 0; i < len; i++) {
			Worksheet sheet = book.getWorksheetAt(i);
			if (sheetId.equals(((SheetCtrl)sheet).getUuid())) {
				spreadsheet.setSelectedSheetDirectly(sheet.getSheetName(), cacheInClient, row, col, 
						left, top, right, bottom,
						highlightLeft, highlightTop, highlightRight, highlightBottom);
				break;
			}
		}
	}
}
