/* SelectionChangeEvent.java

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
import org.zkoss.zss.api.model.Sheet;
/**
 * Event class about selection of cell
 * @author Dennis.Chen
 * @since 3.0.0
 */
public class CellSelectionUpdateEvent extends CellSelectionEvent{

	private static final long serialVersionUID = 1L;

	public enum Action {
		MOVE, //move
		RESIZE //resize
	}
	
	private Action _action;
	private int _origleft;
	private int _origtop;
	private int _origright;
	private int _origbottom;

	public CellSelectionUpdateEvent(String name, Component target,Sheet sheet,SelectionType type, Action action,
			int left, int top,int right, int bottom,
			int origleft, int origtop,int origright, int origbottom) {
		super(name, target, sheet,type,left,top,right,bottom);
		_action = action;
		_origleft = origleft;
		_origtop = origtop;
		_origright = origright;
		_origbottom = origbottom;
	}

	/**
	 * Returns the action of this event.
	 */
	public Action getAction() {
		return _action;
	}
	
	public int getOrigleft() {
		return _origleft;
	}

	public int getOrigtop() {
		return _origtop;
	}

	public int getOrigright() {
		return _origright;
	}

	public int getOrigbottom() {
		return _origbottom;
	}
	
	
}
