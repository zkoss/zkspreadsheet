/* SSWorkbookCtrl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 25, 2010 10:53:10 AM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.image.Image;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.cell.EditHelper;
import org.zkoss.zss.app.event.ExportHelper;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zssex.ui.widget.ImageWidget;
import org.zkoss.zul.Messagebox;

/**
 * @author Sam
 *
 */
public class SSWorkbookCtrl implements WorkbookCtrl {

	private Spreadsheet spreadsheet;
	private Book book;
	
	public SSWorkbookCtrl(Book book, Spreadsheet spreadsheet) {
		this.book = book;
		this.spreadsheet = spreadsheet;
	}
	
	public void clearSelectionContent() {
		CellHelper.clearContent(spreadsheet, SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
	}

	public void clearSelectionStyle() {
		CellHelper.clearStyle(spreadsheet, SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
	}

	public void copySelection() {
		EditHelper.doCopy(spreadsheet);
	}


	public void cutSelection() {
		EditHelper.doCut(spreadsheet);
	}
	
	public void pasteSelection() {
		EditHelper.doPaste(spreadsheet);
	}

	public void insertColumnLeft() {
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireColumnRight(spreadsheet.getSelectedSheet(), 
				rect.getTop(), rect.getLeft());
	}

	public void deleteColumn() {
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireColumnLeft(spreadsheet.getSelectedSheet(), rect.getTop(), rect.getLeft());
	}

	public void insertRowAbove() {
		CellHelper.shiftEntireRowDown(spreadsheet.getSelectedSheet(), 
				spreadsheet.getSelection().getTop(), 
				spreadsheet.getSelection().getLeft());
	}

	public void deleteRow() {
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireRowUp(spreadsheet.getSelectedSheet(), rect.getTop(), rect.getLeft());
	}


	public void setSelectedSheet(String name) {
		spreadsheet.setSelectedSheet(name);
	}

	public void hide(boolean hide) {
		Rect rect = spreadsheet.getSelection();
		Ranges.range(spreadsheet.getSelectedSheet(), 
				rect.getTop(), rect.getLeft(), 
				rect.getBottom(), rect.getRight()).setHidden(hide);
	}

	public void insertImage(Media media) {
		try {
			if (media instanceof org.zkoss.image.Image) {
				ImageWidget image = new ImageWidget();
				image.setContent((Image) media);

				int col = spreadsheet.getSelection().getLeft();
				int row = spreadsheet.getSelection().getTop();
				image.setRow(row);
				image.setColumn(col);
				SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
				ctrl.addWidget(image);
			} else if (media != null) {
				Messagebox.show("Not an image: " + media, "Error",
						Messagebox.OK, Messagebox.ERROR);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}




	public void insertSheet() {
		int sheetCount = spreadsheet.getBook().getNumberOfSheets();
		Sheet addedSheet = spreadsheet.getBook().createSheet("sheet " + (sheetCount + 1));
	}

	public void openExportPdfDialog() {
		ExportHelper.doExportToPDF(spreadsheet);
	}

	public void reGainFocus() {
		Clients.evalJavaScript("zk.Widget.$('" + spreadsheet.getUuid() + "').focus(false);");
	}


	public void renameSelectedSheet(String name) {
		SheetHelper.renameSheet(spreadsheet, name);
	}

	public void shiftCell(int direction) {
		Sheet sheet = spreadsheet.getSelectedSheet();
		Rect rect = spreadsheet.getSelection();
		
		switch (direction) {
		case DesktopWorkbenchContext.SHIFT_CELL_UP:
			CellHelper.shiftCellUp(sheet, 
					rect.getTop(), 
					rect.getLeft());
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_RIGHT:
			CellHelper.shiftCellRight(sheet, 
					rect.getTop(), 
					rect.getLeft());
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_DOWN:
			CellHelper.shiftCellDown(sheet, 
					rect.getTop(), 
					rect.getLeft());
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_LEFT:
			CellHelper.shiftCellLeft(sheet, 
					rect.getTop(), 
					rect.getLeft());
			break;
		}
	}

	public void sort(boolean isSortDescending) {
		if (isSortDescending)
			CellHelper.sortDescending(spreadsheet.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
		else
			CellHelper.sortAscending(spreadsheet.getSelectedSheet(), SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
	}

	public void setColumnFreeze(int columnFreeze) {
		spreadsheet.setColumnfreeze(columnFreeze);
	}

	public void setRowFreeze(int rowfreeze) {
		spreadsheet.setRowfreeze(rowfreeze);
	}

}
