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
import org.zkoss.zss.model.sys.XSheet;

/**
 * Event class about selection of cell
 * @author Dennis.Chen
 */
public class CellSelectionEvent extends Event{
	
	
	public static final int SELECT_CELLS = 0x01;
	public static final int SELECT_ROW = 0x02;
	public static final int SELECT_COLUMN = 0x03;
	public static final int SELECT_ALL = 0x04;
	
	private XSheet _sheet;
	private int _action;
	private int _left;
	private int _top;
	private int _right;
	private int _bottom;

	public CellSelectionEvent(String name, Component target,XSheet sheet,int action, int left, int top,int right, int bottom, Object data) {
		super(name, target, data);
		_sheet = sheet;
		_action = action;
		_left = left;
		_top = top;
		_right = right;
		_bottom = bottom;
	}
	
	public CellSelectionEvent(String name, Component target,XSheet sheet,int action, int left, int top,int right, int bottom) {
		this(name,target,sheet,action,left,top,right,bottom,null);
	}
	
	/**
	 * get Sheet
	 * @return sheet the related sheet 
	 */
	public XSheet getSheet(){
		return _sheet;
	}

	/**
	 * Returns the selection type; it can be either {@link #SELECT_ALL}(Select the whole sheet), 
	 * {@link #SELECT_COLUMN}(select the whole column), {@link #SELECT_ROW}(select the whole row), {@link #SELECT_CELLS}(select a rectangle area).
	 * @return the selection type.
	 */
	public int getSelectionType(){
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
