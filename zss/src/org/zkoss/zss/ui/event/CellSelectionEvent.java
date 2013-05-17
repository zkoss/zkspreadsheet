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
import org.zkoss.zss.api.model.Sheet;

/**
 * Event class about selection of cell
 * @author Dennis.Chen
 */
public class CellSelectionEvent extends Event{
	
//	
//	public static final int SELECT_CELLS = 0x01;
//	public static final int SELECT_ROW = 0x02;
//	public static final int SELECT_COLUMN = 0x03;
//	public static final int SELECT_ALL = 0x04;
	
	public enum SelectionType {
		CELL, //cells
		ROW, //row
		COLUMN, //col
		ALL //all
	}
	
	private Sheet _sheet;
	private SelectionType _type;
	private int _left;
	private int _top;
	private int _right;
	private int _bottom;

	public CellSelectionEvent(String name, Component target,Sheet sheet,SelectionType type, int left, int top,int right, int bottom) {
		super(name, target);
		_sheet = sheet;
		_type = type;
		_left = left;
		_top = top;
		_right = right;
		_bottom = bottom;
	}
	
	/**
	 * get Sheet
	 * @return sheet the related sheet 
	 */
	public Sheet getSheet(){
		return _sheet;
	}

	/**
	 * Returns the selection type
	 * @return the selection type.
	 */
	public SelectionType getType(){
		return _type;
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
