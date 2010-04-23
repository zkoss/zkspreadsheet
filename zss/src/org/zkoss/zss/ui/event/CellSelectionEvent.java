/* CellSelectionEvent.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Dec 19, 2007 2:18:10 PM     2007, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2007 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.event;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.Event;
//import org.zkoss.zss.model.Sheet;
import org.apache.poi.ss.usermodel.*;

/**
 * Event class about selection of cell
 * @author Dennis.Chen
 */
public class CellSelectionEvent extends Event{
	
	
	public static final int SELECT_CELLS = 1;
	public static final int SELECT_ROW = 2;
	public static final int SELECT_COLUMN = 3;
	public static final int SELECT_ALL = 4;
	
	private Sheet _sheet;
	private int _action;
	private int _left;
	private int _top;
	private int _right;
	private int _bottom;

	public CellSelectionEvent(String name, Component target,Sheet sheet,int action, int left, int top,int right, int bottom, Object data) {
		super(name, target, data);
		_sheet = sheet;
		_action = action;
		_left = left;
		_top = top;
		_right = right;
		_bottom = bottom;
	}
	
	public CellSelectionEvent(String name, Component target,Sheet sheet,int action, int left, int top,int right, int bottom) {
		this(name,target,sheet,action,left,top,right,bottom,null);
	}
	
	/**
	 * get Sheet
	 * @return sheet the related sheet 
	 */
	public Sheet getSheet(){
		return _sheet;
	}

	
	public int getAction(){
		return _action;
	}
	
	public int getLeft() {
		return _left;
	}

	public int getTop() {
		return _top;
	}

	public int getRight() {
		return _right;
	}

	public int getBottom() {
		return _bottom;
	}
	
	
	
}
