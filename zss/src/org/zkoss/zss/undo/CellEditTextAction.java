package org.zkoss.zss.undo;

import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.undo.imple.AbstractUndoableAction;

public class CellEditTextAction extends AbstractUndoableAction {

	private String[][] oldEditTexts = null;
	
	private final String editText;
	private final String[][] editTexts;
	
	public CellEditTextAction(Sheet sheet,int row, int column, int lastRow,int lastColumn,String editText){
		super(sheet,row,column,lastRow,lastColumn);
		this.editText = editText;
		editTexts = null;
	}
	public CellEditTextAction(Sheet sheet,int row, int column, int lastRow,int lastColumn,String[][] editTexts){
		super(sheet,row,column,lastRow,lastColumn);
		this.editTexts = editTexts;
		editText = null;
	}
	
	@Override
	public String getLabel() {
		return "Edit Cell Text";
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old text
		oldEditTexts = new String[lastRow-row+1][lastColumn-column+1];
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(sheet,i,j);
				oldEditTexts[i-row][j-column] = r.getCellEditText();
				if(editTexts!=null){
					try{
						r.setCellEditText(editTexts[i][j]);
					}catch(IllegalFormulaException x){};//eat in this mode
				}
			}
		}
		if(editText!=null){
			Range r = Ranges.range(sheet,row,column,lastRow,lastColumn);
			try{
				r.setCellEditText(editText);
			}catch(IllegalFormulaException x){};//eat in this mode
		}
		
	}

	@Override
	public boolean isUndoable() {
		return oldEditTexts!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return oldEditTexts==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		if(oldEditTexts!=null){
			for(int i=row;i<=lastRow;i++){
				for(int j=column;j<=lastColumn;j++){
					Range r = Ranges.range(sheet,i,j);
					r.setCellEditText(oldEditTexts[i-row][j-column]);
				}
			}
			oldEditTexts = null;
		}
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append(getLabel()+"["+row+","+column+","+lastRow+","+lastColumn+"]").append(super.toString());
		return sb.toString();
	}
	
	public Rect getSelection(){
		return new Rect(column,row,lastColumn,lastRow);
	}

}
