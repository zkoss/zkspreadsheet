/* DefaultUndoableActionManager.java

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
package org.zkoss.zss.undo.impl;

import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.ui.Position;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.undo.UndoableAction;
import org.zkoss.zss.undo.UndoableActionManager;
/**
 * default undoable action manager
 * @author dennis
 *
 */
public class DefaultUndoableActionManager implements UndoableActionManager {

	int _maxHistory = 10;
	int _index = -1;
	LinkedList<UndoableAction> _actionHisotry = new LinkedList<UndoableAction>();
	
	Spreadsheet _spreadsheet;
	public DefaultUndoableActionManager(){
		
	}
	public DefaultUndoableActionManager(Spreadsheet spreadsheet){
		this._spreadsheet = spreadsheet;
	}
	
	@Override
	public void doAction(UndoableAction action) {
		action.doAction();
		System.out.println(">>>>>>>>>doAction "+action);
		while(_actionHisotry.size()>_index+1){
			_actionHisotry.removeLast();
		}
		while(_actionHisotry.size()>=_maxHistory){
			_actionHisotry.remove(0);
		}
		_actionHisotry.add(action);
		_index = _actionHisotry.size()-1;
	}
	
	private UndoableAction current(){
		return _index>=0 && _index<_actionHisotry.size()?_actionHisotry.get(_index):null;
	}
	private UndoableAction next(){
		int i = _index+1;
		return i>=0 && i<_actionHisotry.size()?_actionHisotry.get(i):null;
	}

	@Override
	public boolean isUndoable() {
		UndoableAction action = current();
		return action!=null && action.isUndoable();
	}

	@Override
	public String getUndoLabel() {
		UndoableAction action = current();
		return action!=null?action.getLabel():null;
	}

	@Override
	public void undoAction() {
		UndoableAction action = current();
		if(action!=null){
			if(action.isUndoable()){
				System.out.println(">>>>>>>>>undo "+action);
				action.undoAction();
				_index--;
				//to solve performance when select all
				Rect selection = getVisibleRect(action.getUndoSelection());
				if(_spreadsheet!=null && selection!=null){
					_spreadsheet.setSelection(selection);
					_spreadsheet.setCellFocus(new Position(selection.getRow(),selection.getColumn()));
				}
				
			}else{
				clear();
			}
		}
	}

	@Override
	public boolean isRedoable() {
		UndoableAction action = next();
		return action!=null && action.isRedoable();
	}
	
	@Override
	public String getRedoLabel() {
		UndoableAction action = next();
		return action!=null?action.getLabel():null;
	}

	@Override
	public void redoAction() {
		UndoableAction action = next();
		if(action!=null){
			if(action.isRedoable()){
				System.out.println(">>>>>>>>>redo "+action);
				action.doAction();
				_index++;
				//to solve performance when select all
				Rect selection = getVisibleRect(action.getRedoSelection());
				if(_spreadsheet!=null && selection!=null){
					_spreadsheet.setSelection(selection);
					_spreadsheet.setCellFocus(new Position(selection.getRow(),selection.getColumn()));
				}
			}else{
				clear();
			}
		}
	}
	
	private Rect getVisibleRect(Rect rect){
		return new Rect(rect.getLeft(),rect.getTop(),Math.min(rect.getRight(), _spreadsheet.getMaxVisibleRows()),
				Math.min(rect.getBottom(), _spreadsheet.getMaxVisibleColumns()));
	}

	@Override
	public void clear() {
		_index = -1;
		_actionHisotry.clear();
		System.out.println(">>>>>>>>>clear "+this);
	}
	@Override
	public void setMaxHsitorySize(int size) {
		if(size < 1){
			throw new IllegalArgumentException("must large than 1");
		}
		//i simply reset it
		clear();
		_maxHistory = size;
	}

}
