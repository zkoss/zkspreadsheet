package org.zkoss.zss.app;

import org.zkoss.zss.ui.Widget;

public class WidgetPosition {
	Widget widget;
	int col;
	int row;
	
	public WidgetPosition(Widget widget, int col, int row){
		this.widget = widget;
		this.col = col;
		this.row = row;
	}
}
