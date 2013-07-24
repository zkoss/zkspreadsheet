package org.zkoss.zss.undo.imple;

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
		while(_actionHisotry.size()>_index+1){
			_actionHisotry.removeLast();
		}
		if(_maxHistory==_actionHisotry.size()){
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
				action.undoAction();
				_index--;
				
				Rect selection = action.getSelection();
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
				action.doAction();
				_index++;
				
				Rect selection = action.getSelection();
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
	public void clear() {
		_index = -1;
		_actionHisotry = new LinkedList<UndoableAction>();
	}

}
