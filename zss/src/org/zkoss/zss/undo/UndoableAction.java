package org.zkoss.zss.undo;

import org.zkoss.zss.ui.Rect;

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
