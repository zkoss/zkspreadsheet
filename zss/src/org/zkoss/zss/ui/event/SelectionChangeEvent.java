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
//import org.zkoss.zss.model.Sheet;
import org.apache.poi.ss.usermodel.Sheet;
/**
 * Event class about selection of cell
 * @author Dennis.Chen
 */
public class SelectionChangeEvent extends CellSelectionEvent{
	
	
	public static final int MODIFY = 11;
	public static final int MOVE = 12;
	
	private int _origleft;
	private int _origtop;
	private int _origright;
	private int _origbottom;

	public SelectionChangeEvent(String name, Component target,Sheet sheet,int action, 
			int left, int top,int right, int bottom,
			int origleft, int origtop,int origright, int origbottom, Object data) {
		super(name, target, sheet,action,left,top,right,bottom,data);
		_origleft = origleft;
		_origtop = origtop;
		_origright = origright;
		_origbottom = origbottom;
	}
	
	public SelectionChangeEvent(String name, Component target,Sheet sheet, int action,
			int left, int top,int right, int bottom,
			int origleft, int origtop,int origright, int origbottom) {
		this(name,target,sheet,action,left,top,right,bottom,origleft,origtop,origright,origbottom,null);
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
