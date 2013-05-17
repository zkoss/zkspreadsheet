package org.zkoss.zss.ui.event;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.model.Sheet;

public class WidgetUpdateEvent extends Event{
	private static final long serialVersionUID = 1L;
	private Sheet _sheet;
	
	private Action _action;
	
	private SheetAnchor _anchor;
	
	public enum Action {
		MOVE,
		RESIZE
	}
	

	public WidgetUpdateEvent(String name, Component target,Sheet sheet, Action action,Object widgetData,SheetAnchor anchor) {
		super(name, target, widgetData);
		_sheet = sheet;
		_action = action;
		_anchor = anchor;
			
	}
	
	/**
	 * get Sheet
	 * @return sheet the related sheet 
	 */
	public Sheet getSheet(){
		return _sheet;
	}
	
	public SheetAnchor getSheetAnchor(){
		return _anchor;
	}
	
	public Action getAction(){
		return _action;
	}
	
}
