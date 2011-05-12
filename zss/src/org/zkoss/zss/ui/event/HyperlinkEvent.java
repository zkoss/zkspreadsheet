/* HyperlinkEvent.java

	Purpose:
		
	Description:
		
	History:
		Jul 19, 2010 2:08:02 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.ui.event;

import java.util.Map;

import org.zkoss.poi.ss.usermodel.Hyperlink;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zk.au.AuRequest;
import org.zkoss.zk.au.AuRequests;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.MouseEvent;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;

/**
 * Event when end user click on the hyperlink of a cell (used with onHyperlink event).
 * @author henrichen
 */
public class HyperlinkEvent extends CellMouseEvent{
	private String _href;
	private int _type;

	public static HyperlinkEvent getHyperlinkEvent(AuRequest request) {
		final Map data = request.getData();
		final Component comp = request.getComponent();
		String sheetId = (String) data.get("sheetId");
		Worksheet sheet = ((Spreadsheet) comp).getSelectedSheet();
		if (!Utils.getSheetUuid(sheet).equals(sheetId))
			return null;
		
		final String name = request.getCommand();
		final int keys = AuRequests.parseKeys(data);
		return new HyperlinkEvent(name, comp, sheet,
				AuRequests.getInt(data, "row", 0, true),
				AuRequests.getInt(data, "col", 0, true),
				(String) data.get("href"),
				AuRequests.getInt(data, "type", 0, true),
				AuRequests.getInt(data, "x", 0, true),
				AuRequests.getInt(data, "y", 0, true),
				AuRequests.getInt(data, "pageX", 0, true),
				AuRequests.getInt(data, "pageY", 0, true), keys);
	}
	public HyperlinkEvent(String name, Component target, Worksheet sheet, int row ,int col, String href, int type, int x, int y,
			int pageX, int pageY, int keys) {
		super(name, target, x, y, keys, sheet, row, col, pageX, pageY);
		this._href = href;
		this._type = type;
	}

	/**
	 * LINK Reference.
	 * @return URI reference.
	 */
	public String getHref() {
		return _href;
	}

	/**
	 * URL LINK type (can be 1: LINK_URL, 2: LINK_DOCUMENT, 3: LINK_EMAIL, 4: LINK_FILE); see {@link Hyperlink}}.
	 * @return link type
	 */
	public int getType() {
		return _type;
	}
}
