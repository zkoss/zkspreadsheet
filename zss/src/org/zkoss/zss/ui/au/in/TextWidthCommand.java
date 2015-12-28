/* CellFetchCommand.java

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
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.mesg.MZk;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.model.impl.AbstractCellAdv;

/**
 * A Command (client to server) for pass cell text width back
 * @author Henri Chen
 * @since 3.8.3
 */
public class TextWidthCommand extends AbstractCommand implements Command {
	public final static String Command = "onZSSTextWidth";

	//-- super --//
	public void process(AuRequest request) {
		final Component comp = request.getComponent();
		if (comp == null)
			throw new UiException(MZk.ILLEGAL_REQUEST_COMPONENT_REQUIRED, TextWidthCommand.class.getCanonicalName());

		final Map data = (Map) request.getData();
		if (data == null || data.size() != 4)
			throw new UiException(MZk.ILLEGAL_REQUEST_WRONG_DATA, new Object[] {Objects.toString(data), TextWidthCommand.class.getCanonicalName() });
		
		String sheetId = (String) data.get("sheetId");
		Sheet sheet = ((Spreadsheet) comp).getSelectedSheet();
		if (!getSheetUuid(sheet).equals(sheetId))
			return;

		int row = (Integer) data.get("row");
		int col = (Integer) data.get("col");
		int width = (Integer) data.get("width");
		
		((AbstractCellAdv)sheet.getInternalSheet().getCell(row, col)).setTextWidth(width);
	}
}
