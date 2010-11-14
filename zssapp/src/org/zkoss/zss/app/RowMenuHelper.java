package org.zkoss.zss.app;

import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.UiException;
//import org.zkoss.zss.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.impl.Utils;
import org.zkoss.zul.Window;

public class RowMenuHelper {
	Spreadsheet spreadsheet;
	boolean isCut = false;// true:cut, false:copy
	Sheet remSheet;

	public RowMenuHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
	}

	public void dispatcher(String type) {
		if (type.equalsIgnoreCase("Insert")) {
			onRowHeaderMenuInsert();
		} else if (type.equalsIgnoreCase("Delete")) {
			onRowHeaderMenuDelete();
		} else if (type.equalsIgnoreCase("Clear")) {
			onRowHeaderMenuClear();
		} else if (type.equalsIgnoreCase("Height")) {
			onRowHeaderMenuHeight();
		} else if (type.equalsIgnoreCase("Format")) {
			//onRowHeaderMenuFormat();
			throw new UiException("Not implement yet");
			// }else if(type.equalsIgnoreCase("Hide")){
			// onRowHeaderMenuHide();
			// }else if(type.equalsIgnoreCase("Unhide")){
			// onRowHeaderMenuUnhidet();
		} else {
			System.out.println("wh not supported " + type);
		}
	}

	public void onRowHeaderMenuDelete() {
//TODO delete rows
throw new UiException("delete rows not implemented yet");
/*
		try {
			
			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			int colSize = spreadsheet.getMaxcolumns();

			if (0 == left && (right + 1) == colSize) {
				Sheet sheet=spreadsheet.getSelectedSheet();
//TODO undo/redo				
//				spreadsheet.pushDeleteRowColState(-1, top, -1, bottom);
//TODO delete rows				
//				sheet.deleteRows(top, bottom);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
*/	}

	public void onRowHeaderMenuInsert() {
//TODO insert rows
throw new UiException("insert rows not implmented yet");
/*		try {
			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			int colSize = spreadsheet.getMaxcolumns();
			
			if (0 == left && (right + 1) == colSize) {
				spreadsheet.pushInsertRowColState(-1, top, -1, bottom);
				spreadsheet.getSelectedSheet().insertRows(top, bottom);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
*/		 
	}

	public void onRowHeaderMenuClear() {
//TODO clear rows
throw new UiException("clear rows not implmented yet");
/*
		try {
			int left = spreadsheet.getSelection().getLeft();
			int right = spreadsheet.getSelection().getRight();
			int top = spreadsheet.getSelection().getTop();
			int bottom = spreadsheet.getSelection().getBottom();
			int colSize = spreadsheet.getMaxcolumns();
			if (0 == left && (right + 1) == colSize) {
				// by whkuo, delete then insert ==> clear?
				// TODO: use Spreadsheet.ExtraCtrl.removeRows is a better idea?
				spreadsheet.getSelectedSheet().deleteRows(top, bottom);
				spreadsheet.getSelectedSheet().insertRows(top, bottom);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
*/	}


	public void onRowHeaderMenuHeight() {
		Window win = (Window) Executions.createComponents(
				"/menus/header_cell/menuHeight.zul", null, null);
		try {
			win.doModal();
			onSetRowHeight();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onSetRowHeight() {
		String heightOK = (String) spreadsheet.getDesktop().getSession()
				.getAttribute("heightOK");
		String heightValue = (String) spreadsheet.getDesktop().getSession()
				.getAttribute("heightValue");
		try {
			int validHeight = Integer.parseInt(heightValue);
			if (heightOK.equals("true")) {
				int left = spreadsheet.getSelection().getLeft();
				int right = spreadsheet.getSelection().getRight();
				int top = spreadsheet.getSelection().getTop();
				int bottom = spreadsheet.getSelection().getBottom();
//TODO undo/redo				
//				spreadsheet.pushRowColSizeState(-1, top, -1, bottom);
				
				int colSize = spreadsheet.getMaxcolumns();
				if (0 == left && (right + 1) == colSize) {
					for (int i = top; i <= bottom; i++) {
						Row row = Utils.getOrCreateRow(spreadsheet.getSelectedSheet(),i);
						row.setHeightInPoints(validHeight);
					}
				}
				spreadsheet.updateFocus(-1, top, -1, bottom);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

//	public void onRowHeaderMenuFormat() {
//		MainWindowCtrl.getInstance().onFormatPopup();
//	}
}