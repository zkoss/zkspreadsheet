package org.zkoss.zss.app;

import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.UiException;
//import org.zkoss.zss.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Window;
import org.apache.poi.ss.usermodel.Sheet;

public class ColumnMenuHelper {
	Spreadsheet spreadsheet;

	boolean isCut = false;// true:cut, false:copy
	Sheet remSheet;
	
	public ColumnMenuHelper(Spreadsheet spreadsheet){
		this.spreadsheet = spreadsheet;
	}
	
	public void dispatcher(String type){
	    if(type.equalsIgnoreCase("Insert")){
			onColumnHeaderMenuInsert();
		}else if(type.equalsIgnoreCase("Delete")){
			onColumnHeaderMenuDelete();
		}if(type.equalsIgnoreCase("Width")){
			onColumnHeaderMenuWidth();
		}else if(type.equalsIgnoreCase("Format")){
			onColumnHeaderMenuFormat();
		}else{
			System.out.println("wh not supported " + type);
		}			
	}
	
	public void onColumnHeaderMenuDelete() {
		try {
			
			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			int rowSize = spreadsheet.getMaxrows();
			
			// have to know if selected whole column, or just a partial rect.
			// System.out.println(spreadsheet.getSelection().toString());
			// by whkuo, it's strange to get rowSize from spreadsheet, not the
			// sheet.
			// int rowSize = spreadsheet.getSelectedSheet().getRowSize();
			// System.out.println("rowsize" + String.valueOf(rowSize));
			if (0 == top && (bottom + 1) == rowSize) {
//TODO undo/redo				
//				spreadsheet.pushDeleteRowColState(left, -1, right, -1);
				
				/*Sheet sheet=spreadsheet.getSelectedSheet();
				for(int row=0; row<=rowSize; row++)
					for(int col=left; col<right; col++){
						((SheetImpl)sheet).cleanDependRef(row, col);
					}
				*/
//TODO deleteColumns				
//				spreadsheet.getSelectedSheet().deleteColumns(left, right);
				/**
				 * Sam. remove concurrent function
				 */
				//spreadsheet.notifyOperation(Operation.DELETE_ROW_COL);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
//TODO remove me		
throw new UiException("delete columns is not implemented yet");
	}

	public void onColumnHeaderMenuInsert() {
		try {
			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			int rowSize = spreadsheet.getMaxrows();
			
			if (0 == top && (bottom + 1) == rowSize) {
//TODO undo/redo				
//				spreadsheet.pushInsertRowColState(left, -1, right, -1);
//TODO insertColumns
//				spreadsheet.getSelectedSheet().insertColumns(left, right);
				/**
				 * Sam remove concurrent operation
				 */
				//spreadsheet.notifyOperation(Operation.INSERT_ROW_COL);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
//TODO remove me		
throw new UiException("insert columns is not implemented yet");
	}


	public void onColumnHeaderMenuWidth() {
		
		Window win = (Window) Executions.createComponents("/menus/header_cell/menuWidth.zul",
				null, null);
		try {
			win.doModal();
			onSetColumnWidth();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	public void onSetColumnWidth() {		
		String widthOK = (String)spreadsheet.getDesktop().getSession().getAttribute(
				"widthOK");
		String widthValue = (String)spreadsheet.getDesktop().getSession().getAttribute(
				"widthValue");
		try {
			int validWidth = Integer.parseInt(widthValue);
			if (widthOK.equals("true")) {				
				int left = spreadsheet.getSelection().getLeft();
				int right = spreadsheet.getSelection().getRight();
				int top = spreadsheet.getSelection().getTop();
				int bottom = spreadsheet.getSelection().getBottom();

//TODO undo/redo
//				spreadsheet.pushRowColSizeState(left, -1, right, -1);
				
				int rowSize = spreadsheet.getMaxrows();
				if (0 == top && (bottom + 1) == rowSize) {
					for(int i = left;i<=right;i++){
						spreadsheet.getSelectedSheet().setColumnWidth(i, validWidth);
					}					
				}
				/**
				 * Sam remove concurrent function
				 */
				//spreadsheet.notifyOperation(Operation.ROW_COL_SIZE_CHANGED);
				spreadsheet.updateFocus(left, -1, right, -1);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	public void onColumnHeaderMenuFormat(){
		MainWindowCtrl.getInstance().onFormatPopup();
	}
	
	
}
