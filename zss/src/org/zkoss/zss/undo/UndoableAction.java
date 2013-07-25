/* UndoableAction.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/7/25, Created by Dennis.Chen
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.undo;

import org.zkoss.zss.ui.Rect;
/**
 * 
 * @author dennis
 *
 */
public interface UndoableAction{

	/**
	 * 
	 * @return the label of this action
	 */
	public String getLabel();
	
	/**
	 * do the action, either first time or redo
	 */
	public void doAction();
	
	/**
	 * Check if still undoable or not
	 * @return
	 */
	public boolean isUndoable();
	
	/**
	 * Check if still redoable or not
	 * @return
	 */
	public boolean isRedoable();
	
	/**
	 * Undo the action
	 */
	public void undoAction();
	
	/**
	 * 
	 * @return Selection of this action, null if doesn't provided;
	 */
	public Rect getSelection();
}
