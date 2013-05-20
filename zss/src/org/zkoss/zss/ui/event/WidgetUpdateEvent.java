/* WidgetUpdateEvent.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/20 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.event;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.model.Sheet;

/**
 * Event about widget update
 * @author dennis
 * @since 3.0.0
 */
public class WidgetUpdateEvent extends Event{
	private static final long serialVersionUID = 1L;
	private Sheet _sheet;
	
	private WidgetAction _action;
	
	private SheetAnchor _anchor;

	

	public WidgetUpdateEvent(String name, Component target,Sheet sheet, WidgetAction action,Object widgetData,SheetAnchor anchor) {
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
	
	public WidgetAction getAction(){
		return _action;
	}
	
}
