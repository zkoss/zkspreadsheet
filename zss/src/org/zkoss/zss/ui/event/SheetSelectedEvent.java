package org.zkoss.zss.ui.event;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.api.model.Sheet;

public class SheetSelectedEvent extends Event{
	private static final long serialVersionUID = 1L;
	private Sheet _sheet;
	private Sheet _previousSheet;
	
	
	public SheetSelectedEvent(String name, Component target,Sheet sheet,Sheet previousSheet) {
		super(name, target, sheet);
		_sheet = sheet;
		_previousSheet = previousSheet;
	}
	
	public Sheet getSheet(){
		return _sheet;
	}
	
	public Sheet getPreviousSheet(){
		return _previousSheet;
	}

}
