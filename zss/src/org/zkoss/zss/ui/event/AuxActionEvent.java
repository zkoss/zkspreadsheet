/* AuxActionEvent.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/05/09 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.event;

import java.util.Map;

import org.zkoss.zk.ui.Component;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;

/**
 * @author dennis
 *
 */
public class AuxActionEvent extends org.zkoss.zk.ui.event.Event {

	private static final long serialVersionUID = 1L;
	final Sheet sheet;
	final String action;
	final Rect selection;
	
	public AuxActionEvent(String name, Component target, Sheet sheet, String action,Rect selection, Map data) {
		super(name, target, data);
		this.sheet = sheet;
		this.action = action;
		this.selection = selection;
	}

	public Sheet getSheet() {
		return sheet;
	}

	public String getAction() {
		return action;
	}

	public Rect getSelection() {
		return selection;
	}

	public Map getExtraData() {
		return (Map)getData();
	}
	
}
