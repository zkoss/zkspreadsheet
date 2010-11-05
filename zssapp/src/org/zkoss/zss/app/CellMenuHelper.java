package org.zkoss.zss.app;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.zkoss.image.Image;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Path;
//import org.zkoss.zss.model.Sheet;
//import org.zkoss.zss.model.impl.RangeSimple;
//import org.zkoss.zss.model.impl.SheetImpl;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zssex.ui.widget.ImageWidget;
import org.zkoss.zul.Fileupload;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Window;


public class CellMenuHelper {
	Spreadsheet spreadsheet;
	boolean isCut = false;// true:cut, false:copy
	Sheet remSheet;

	public CellMenuHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
	}

	public void dispatcher(String type) {
		if (type.equalsIgnoreCase("InsertPic")) {
			onCellMenuInsertPic();
		} else if (type.equalsIgnoreCase("RemovePic")) {
			onCellMenuRemovePic();
		} else if (type.equalsIgnoreCase("RangeAdd")) {
			onCellMenuRangeAdd();
		} else if (type.equalsIgnoreCase("RangeDelete")) {
			onCellMenuRangeDelete();
		} else if (type.equalsIgnoreCase("RangeChoose")) {
			onCellMenuRangeChoose();
		} else {
			System.out.println("CellMenuHelper type not supported : " + type);
		}
	}


	List widgetList = new ArrayList();

	/**
	 * Sam. try to remove this method
	 */
	public void onCellMenuInsertPic() {
		try {
			Object media = Fileupload.get();
			if (media instanceof org.zkoss.image.Image) {
				System.out.println("instanceof Image");
				ImageWidget image = new ImageWidget();
				image.setContent((Image) media);

				int col = spreadsheet.getSelection().getLeft();
				int row = spreadsheet.getSelection().getTop();
				System.out.println("row: " + row + ", col: " + col);
				image.setRow(row);
				image.setColumn(col);
				WidgetPosition widgetPosition = new WidgetPosition(image, col, row);
				widgetList.add(widgetPosition);

				SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
				ctrl.addWidget(image);
				/**
				 * Sam. try it
				 */
				//TODO: should not need to call invalidate(), post bug
				//spreadsheet.invalidate();
			} else if (media != null) {
				Messagebox.show("Not an image: " + media, "Error", Messagebox.OK, Messagebox.ERROR);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onCellMenuRemovePic() {
		try {
			int col = spreadsheet.getSelection().getLeft();
			int row = spreadsheet.getSelection().getTop();
			Iterator iter = widgetList.iterator(); 
			for (;iter.hasNext();) {
				WidgetPosition wp = (WidgetPosition) iter.next();
				if (wp.col == col && wp.row == row) {
					SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
					ctrl.removeWidget(wp.widget);
					widgetList.remove(wp);
					spreadsheet.invalidate();
					break;
				}
			}
		} catch (Exception e) {
			// TODO:
			e.printStackTrace();
		}
	}

	public void onCellMenuRangeAdd() {
		Window win = (Window) Executions.createComponents("/menus/range/rangeListAdd.zul", null, null);
		try {
			win.doModal();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onCellMenuRangeDelete() {
		Window win = (Window) Executions.createComponents("/menus/range/rangeListDelete.zul", null, null);
		try {
			win.doModal();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void onCellMenuRangeChoose() {
		Window win = (Window) Executions.createComponents("/menus/range/rangeListChoose.zul", null, null);
		try {
			win.doModal();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
