package org.zkoss.zss.ui.event;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.api.model.Sheet;

public class WidgetEvent extends Event{
	private static final long serialVersionUID = 1L;
	private Sheet _sheet;

	public WidgetEvent(String name, Component target,Sheet sheet, Object data) {
		super(name, target, data);
		_sheet = sheet;
	}
	
	/**
	 * get Sheet
	 * @return sheet the related sheet 
	 */
	public Sheet getSheet(){
		return _sheet;
	}
	
}
