/* HeaderEvent.java

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
import org.apache.poi.ss.usermodel.Sheet;

/**
 * A class from handle event which about header
 * @author Dennis.Chen
 */
public class HeaderEvent extends Event{
	
	static public final int TOP_HEADER = 0;
	static public final int LEFT_HEADER = 1;
	
	private Sheet _sheet;
	private int _type;
	private int _index;
	private boolean _hidden;

	public HeaderEvent(String name, Component target,Sheet sheet, int type ,int index,Object data, boolean hidden) {
		super(name, target, data);
		_sheet = sheet;
		this._type = type;
		this._index = index;
		this._hidden = hidden;
	}
	
	public HeaderEvent(String name, Component target,Sheet sheet, int type ,int index, boolean hidden) {
		this(name,target,sheet,type,index,null,hidden);
	}
	
	/**
	 * get Sheet
	 * @return sheet 
	 */
	public Sheet getSheet(){
		return _sheet;
	}
	
	/**
	 * get index of the header, if the {@link #getType} return @link HeaderEvent#TOP_HEADER} then it is column index, otherwise it is row index 
	 * @return row index
	 */
	public int getIndex(){
		return _index;
	}
	
	
	/**
	 * get type of this event, it will be {@link HeaderEvent#TOP_HEADER} or (@link HeaderEvent#LEFT_HEADER} 
	 * @return the type of header
	 */
	public int getType(){
		return _type;
	}
	
	/**
	 * Returns whether request hidden(true)/unhidden(false) of this column/row.
	 * @return whether request hidden(true)/unhidden(false) of this column/row.
	 */
	public boolean isHidden() {
		return _hidden;
	}
	
	public String toString(){
		StringBuffer sb = new StringBuffer();
		sb.append(super.toString());
		sb.append("[").append("type:").append(_type).append(",index:").append(_index).append(",data:").append(getData()).append(",hidden:").append(isHidden()).append("]");
		return sb.toString();
	}

}
