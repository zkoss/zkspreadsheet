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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.zkoss.image.Image;
import org.zkoss.poi.ss.usermodel.Sheet;
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zss.app.cell.CellHelper;
import org.zkoss.zss.app.cell.EditHelper;
import org.zkoss.zss.app.event.ExportHelper;
import org.zkoss.zss.app.sheet.SheetHelper;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.ui.Rect;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.Widget;
import org.zkoss.zss.ui.impl.Utils;
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
	
	private String lastSheetName = null;
	private HashMap<String, List<Widget>> sheetWidgets = new HashMap<String, List<Widget>>(); 
	
	public SSWorkbookCtrl(Book book, Spreadsheet spreadsheet) {
		this.book = book;
		this.spreadsheet = spreadsheet;
	}
	
	public void clearSelectionContent() {
		if (spreadsheet.getSelection() == null)
			return;
		
		CellHelper.clearContent(spreadsheet, SheetHelper.getSpreadsheetMaxSelection(spreadsheet));
	}

	public void clearSelectionStyle() {
		if (spreadsheet.getSelection() == null)
			return;
		
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
		if (spreadsheet.getSelection() == null)
			return;
		
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireColumnRight(spreadsheet.getSelectedSheet(), 
				rect.getLeft(), rect.getRight());
	}

	public void deleteColumn() {
		Rect rect = spreadsheet.getSelection();
		if (rect == null)
			return;
		
		CellHelper.shiftEntireColumnLeft(spreadsheet.getSelectedSheet(), 
				rect.getLeft(), rect.getRight());
	}

	public void insertRowAbove() {
		CellHelper.shiftEntireRowDown(spreadsheet.getSelectedSheet(), 
				spreadsheet.getSelection().getTop(), 
				spreadsheet.getSelection().getBottom());
	}

	public void deleteRow() {
		Rect rect = spreadsheet.getSelection();
		CellHelper.shiftEntireRowUp(spreadsheet.getSelectedSheet(), rect.getTop(), rect.getBottom());
	}


	public void setSelectedSheet(String name) {
		//TODO: remove last sheet widget shall not handle by AP
		Sheet lastsheet = spreadsheet.getSelectedSheet();
		List<Widget> rmWgtList = sheetWidgets.get(lastSheetName);
		if (rmWgtList != null) {
			SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
			for (Widget w : rmWgtList) {
				ctrl.removeWidget(w);
			}
		}
		Integer lastIdx = Integer.valueOf(book.getSheetIndex(lastsheet));
		if (lastIdx < 0) //sheet deleted
			sheetWidgets.remove(lastSheetName);
		
		spreadsheet.setSelectedSheet(name);		
		//handle the copy/cut highlight
		final Sheet sheet = EditHelper.getSourceSheet(spreadsheet);
		if (sheet != null) {
			if (sheet.equals(spreadsheet.getSelectedSheet())) {
				spreadsheet.setHighlight(EditHelper.getSourceRange(spreadsheet));
			} else {
				spreadsheet.setHighlight(null);
			}
		}
		
		//TODO: insert sheet widget shall not handle by AP
		lastSheetName = name;
		List<Widget> addWgtList = sheetWidgets.get(lastSheetName);
		if (addWgtList != null) {
			SpreadsheetCtrl ctrl = (SpreadsheetCtrl) spreadsheet.getExtraCtrl();
			for (Widget w : addWgtList) {
				ctrl.addWidget(w);
			}
		}
	}

	public void hide(boolean hide) {
		Rect rect = spreadsheet.getSelection();
		if (rect == null)
			return;
		
		Ranges.range(spreadsheet.getSelectedSheet(), 
				rect.getTop(), rect.getLeft(), 
				rect.getBottom(), rect.getRight()).setHidden(hide);
	}

	public void insertImage(Media media) {
		//TODO: insert image shall not handle by AP
		try {
			if (media instanceof org.zkoss.image.Image) {
				ImageWidget image = new ImageWidget();
				image.setContent((Image) media);
				image.setRow(spreadsheet.getSelection().getTop());
				image.setColumn(spreadsheet.getSelection().getLeft());
				
				Sheet seldSheet = spreadsheet.getSelectedSheet();
				String sheetName = seldSheet.getSheetName();
				List<Widget> wgtList = sheetWidgets.get(sheetName);
				if (wgtList == null)
					sheetWidgets.put(sheetName, wgtList = new ArrayList<Widget>());
				wgtList.add(image);
				
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
		spreadsheet.getBook().createSheet("sheet " + (sheetCount + 1));
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
			CellHelper.shiftCellUp(sheet, rect); 
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_RIGHT:
			CellHelper.shiftCellRight(sheet, rect);
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_DOWN:
			CellHelper.shiftCellDown(sheet, rect); 
			break;
		case DesktopWorkbenchContext.SHIFT_CELL_LEFT:
			CellHelper.shiftCellLeft(sheet, rect); 
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

	public void insertFormula(String formula) {
		Rect rect = spreadsheet.getSelection();
		Range rng = Ranges.range(spreadsheet.getSelectedSheet(), rect.getTop(), rect.getLeft());
		//Note. can not catch evaluate exception here
		rng.setEditText(formula);
	}

	public void addEventListener(String evtnm, EventListener listener) {
		spreadsheet.addEventListener(evtnm, listener);
	}
	
	public boolean removeEventListener(String evtnm, EventListener listener) {
		return spreadsheet.removeEventListener(evtnm, listener);
	}

	public String getCurrentCellPosition() {
		return (String)spreadsheet.getColumntitle(spreadsheet.getSelection().getLeft()) +
			(String)spreadsheet.getRowtitle(spreadsheet.getSelection().getTop());
	}

	public void setDataFormat(String format) {
		Utils.setDataFormat(spreadsheet.getSelectedSheet(), 
				spreadsheet.getSelection(), format);
	}

	public void save() {
		throw new UiException("save not implement yet");
	}
}