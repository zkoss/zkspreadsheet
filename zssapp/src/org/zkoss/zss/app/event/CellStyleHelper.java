/* CellStyleHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 16, 2010 9:16:24 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.app.event;

import org.zkoss.poi.ss.usermodel.BorderStyle;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zss.app.MainWindowCtrl;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;

/**
 * Spreadsheet helper for set border, font, background event.
 * @author Sam
 *
 */
public final class CellStyleHelper {
	private CellStyleHelper(){}
	
	/**
	 * Execute border base on event's parameter
	 * @param event
	 */
	public static void onBorderEventHandler(ForwardEvent event) {
		String param=(String)event.getData();
		if (param == null)
			return;
		
		MainWindowCtrl ctrl = MainWindowCtrl.getInstance();
		Spreadsheet ss = ctrl.getSpreadsheet();
		
		//not permit to changing color and style
		final BorderStyle lineStyle = BorderStyle.MEDIUM;
		final String color = "#000000";

		if (param.equals(Labels.getLabel("border.no")))
			Utils.setBorder(ss.getSelectedSheet(), ss.getSelection(), BookHelper.BORDER_FULL, BorderStyle.NONE, color);
		else {
			Utils.setBorder(ss.getSelectedSheet(), ss.getSelection(), getBorderType(param), BorderStyle.MEDIUM, color);
		}
	}
	
	/**
	 * Returns the border type base on i3-label
	 * <p> Default: returns {@link #BookHelper.BORDER_EDGE_BOTTOM}, if no match.
	 * @param i3label
	 */
	public static short getBorderType(String i3label) {
		if (i3label == null || i3label.equals(Labels.getLabel("border.bottom"))) {
			return BookHelper.BORDER_EDGE_BOTTOM;
		}
		
		if (i3label.equals(Labels.getLabel("border.bottom")))
			return BookHelper.BORDER_EDGE_BOTTOM;
		else if (i3label.equals(Labels.getLabel("border.right")))
			return BookHelper.BORDER_EDGE_RIGHT;
		else if (i3label.equals(Labels.getLabel("border.top")))
			return BookHelper.BORDER_EDGE_TOP;
		else if (i3label.equals(Labels.getLabel("border.left")))
			return BookHelper.BORDER_EDGE_LEFT;
		else if (i3label.equals(Labels.getLabel("border.insideHorizontal")))
			return BookHelper.BORDER_INSIDE_HORIZONTAL;
		else if (i3label.equals(Labels.getLabel("border.insideVertical")))
			return BookHelper.BORDER_INSIDE_VERTICAL;
		else if (i3label.equals(Labels.getLabel("border.diagonalDown")))
			return BookHelper.BORDER_DIAGONAL_DOWN;
		else if (i3label.equals(Labels.getLabel("border.diagonalUp")))
			return BookHelper.BORDER_DIAGONAL_UP;
		else if (i3label.equals(Labels.getLabel("border.full")))
			return BookHelper.BORDER_FULL;
		else if (i3label.equals(Labels.getLabel("border.outside")))
			return BookHelper.BORDER_OUTLINE;
		else if (i3label.equals(Labels.getLabel("border.inside")))
			return BookHelper.BORDER_INSIDE;
		else if (i3label.equals(Labels.getLabel("border.diagonal")))
			return BookHelper.BORDER_DIAGONAL;

		return BookHelper.BORDER_EDGE_BOTTOM;
	}	
}