package org.zkoss.zss.undo.imple;

import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.undo.UndoableAction;

abstract public class AbstractUndoableAction implements UndoableAction {

	protected final Sheet sheet;
	protected final int row,column,lastRow,lastColumn;
	public AbstractUndoableAction(Sheet sheet,int row, int column, int lastRow,int lastColumn){
		this.sheet = sheet;
		this.row = row;
		this.column = column;
		this.lastRow = lastRow;
		this.lastColumn = lastColumn;
	}

	protected boolean isSheetAvailable(){
		try{
			Book book = sheet.getBook();
			return book.getSheetIndex(sheet)>=0;
		}catch(Exception x){}
		return false;
	}
	
	protected boolean isSheetProtected(){
		return sheet.isProtected();
	}

}
