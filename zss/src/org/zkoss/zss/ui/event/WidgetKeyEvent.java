package org.zkoss.zss.ui.event;

import org.zkoss.zk.ui.Component;
import org.zkoss.zss.api.model.Sheet;

public class WidgetKeyEvent extends org.zkoss.zk.ui.event.KeyEvent{
	private static final long serialVersionUID = 1L;
	private Sheet _sheet;
	private Object _data;

	public WidgetKeyEvent(String name, Component target, Sheet sheet, Object data,int keyCode,
			boolean ctrlKey, boolean shiftKey, boolean altKey) {
		super(name,target,keyCode,ctrlKey,shiftKey,altKey);
		_sheet = sheet;
		_data = data;
	}
	
	/**
	 * get Sheet
	 * @return sheet the related sheet 
	 */
	public Sheet getSheet(){
		return _sheet;
	}
	
	@Override
	public Object getData(){
		return _data;
	}
	
}
