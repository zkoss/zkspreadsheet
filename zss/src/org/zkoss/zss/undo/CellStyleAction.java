package org.zkoss.zss.undo;


import org.zkoss.zss.api.CellOperationUtil;
import org.zkoss.zss.api.CellOperationUtil.CellStyleApplier;
import org.zkoss.zss.api.IllegalFormulaException;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.undo.imple.AbstractUndoableAction;

public class CellStyleAction extends AbstractUndoableAction {

	private CellStyle[][] oldStyles = null;
	
	private final CellStyleApplier styleApplier;
	
	
	public CellStyleAction(Sheet sheet,int row, int column, int lastRow,int lastColumn,CellStyleApplier styleApplier){
		super(sheet,row,column,lastRow,lastColumn);
		this.styleApplier = styleApplier;
	}
	
	@Override
	public String getLabel() {
		return "Edit Cell Style";
	}

	@Override
	public void doAction() {
		if(isSheetProtected()) return;
		//keep old text
		oldStyles = new CellStyle[lastRow-row+1][lastColumn-column+1];
		for(int i=row;i<=lastRow;i++){
			for(int j=column;j<=lastColumn;j++){
				Range r = Ranges.range(sheet,i,j);
				oldStyles[i-row][j-column] = r.getCellStyle();
			}
		}
		Range r = Ranges.range(sheet,row,column,lastRow,lastColumn);
		try{
			CellOperationUtil.applyCellStyle(r, styleApplier);
		}catch(IllegalFormulaException x){};//eat in this mode
	}

	@Override
	public boolean isUndoable() {
		return oldStyles!=null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public boolean isRedoable() {
		return oldStyles==null && isSheetAvailable() && !isSheetProtected();
	}

	@Override
	public void undoAction() {
		if(isSheetProtected()) return;
		if(oldStyles!=null){
			for(int i=row;i<=lastRow;i++){
				for(int j=column;j<=lastColumn;j++){
					Range r = Ranges.range(sheet,i,j);
					r.setCellStyle(oldStyles[i-row][j-column]);
				}
			}
			oldStyles = null;
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
