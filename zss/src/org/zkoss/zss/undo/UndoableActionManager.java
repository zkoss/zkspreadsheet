package org.zkoss.zss.undo;


public interface UndoableActionManager {
	
	public void doAction(UndoableAction action);
	public boolean isUndoable();
	public String getUndoLabel();
	public void undoAction();
	
	public boolean isRedoable();
	public String getRedoLabel();
	public void redoAction();
	
	public void clear();
}
