/* DesktopSheetContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2010 2:51:14 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Desktop;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zss.app.Consts;

/**
 * @author Sam
 *
 */
public class DesktopSheetContext extends AbstractBaseContext{

	public final static int SHIFT_CELL_UP = 0;
	public final static int SHIFT_CELL_RIGHT = 1;
	public final static int SHIFT_CELL_DOWN = 2;
	public final static int SHIFT_CELL_LEFT = 3;
	
	public static DesktopSheetContext getInstance(Desktop desktop) {
		DesktopSheetContext ctrl = 
			(DesktopSheetContext) desktop.getAttribute("DesktopSheetContext");
		if(ctrl==null){
			desktop.setAttribute("DesktopSheetContext", 
					ctrl = new DesktopSheetContext());
		}
		return ctrl;
	}

	/* View */
	public void reGainFocus() {
		listenerStore.fire(new Event(Consts.ON_SHEET_REGAIN_FOCUS));
	}

	public void refresh() {
		listenerStore.fire(new Event(Consts.ON_SHEET_REFRESH));
	}
	
	public void renameSelectedSheet(String name) {
		listenerStore.fire(new Event(Consts.ON_SHEET_RENAME, null, name));
	}
	
	//TODO: remove, change to use main controller to set sheet
	public void fireSheetSelected(String name) {
		listenerStore.fire(new Event(Consts.ON_SHEET_SELECT, null, name));
	}

	public void doSheetOpen(boolean open) {
		listenerStore.fire(new Event(Consts.ON_SHEET_OPEN, null, Boolean.valueOf(open)));
	}
	
	public void insertSheet() {
		listenerStore.fire(new Event(Consts.ON_SHEET_INSERT));
	}
	
	public void insertImage(Media media) {
		listenerStore.fire(new Event(Consts.ON_SHEET_INSERT_IMAGE, null, media));
	}
	
	/* Edit */
	public void cutSelection() {
		listenerStore.fire(new Event(Consts.ON_SHEET_CUT_SELECTION));
	}
	
	public void copySelection() {
		listenerStore.fire(new Event(Consts.ON_SHEET_COPY_SELECTION));
	}
	
	public void pasteSelection() {
		listenerStore.fire(new Event(Consts.ON_SHEET_PASTE_SELECTION));
	}
	
	public void clearSelectionContent() {
		listenerStore.fire(new Event(Consts.ON_SHEET_CLEAR_SELECTION_CONTENT));
	}

	public void clearSelectionStyle() {
		listenerStore.fire(new Event(Consts.ON_SHEET_CLEAR_SELECTION_STYLE));
	}
	public void insertRow() {
		listenerStore.fire(new Event(Consts.ON_SHEET_INSERT_ROW));
	}
	public void deleteRow() {
		listenerStore.fire(new Event(Consts.ON_SHEET_DELETE_ROW));
	}
	public void insertColumn() {
		listenerStore.fire(new Event(Consts.ON_SHEET_INSERT_COLUMN));
	}
	public void deleteColumn() {
		listenerStore.fire(new Event(Consts.ON_SHEET_DELETE_COLUMN));
	}
	public void hide(boolean hide) {
		listenerStore.fire(new Event(Consts.ON_SHEET_HIDE, null, Boolean.valueOf(hide)));
	}
	public void sort(boolean isSortDescending) {
		listenerStore.fire(new Event(Consts.ON_SHEET_SORT, null, Boolean.valueOf(isSortDescending)));
	}
	public void shiftCell(int direction) {
		listenerStore.fire(new Event(Consts.ON_SHEET_SHIFT_CELL, null, Integer.valueOf(direction)));
	}
	public void mergeCell() {
		listenerStore.fire(new Event(Consts.ON_SHEET_MERGE_CELL));
	}
	public void insertFormula(String formula) {
		listenerStore.fire(new Event(Consts.ON_SHEET_INSERT_FORMULA, null, formula));
	}
	
	/* Dialog */
	//TODO: shall remove open dialog, context shall not know about UI
	public void openInsertFormulaDialog() {
		listenerStore.fire(new Event(Consts.ON_SHEET_INSERT_FORMULZ_DIALOG));
	}
	public void openExportPdfDialog() {
		listenerStore.fire(new Event(Consts.ON_SHEET_EXPORT_PDF_DIALOG));
	}
	public void openPasteSpecialDialog() {
		listenerStore.fire(new Event(Consts.ON_SHEET_PASTE_SPECIAL_DIALOG));
	}
	public void openModifyRowHeightDialog() {
		listenerStore.fire(new Event(Consts.ON_SHEET_MODIFY_ROW_HEIGHT_DIALOG));
	}
	public void openCustomSortDialog() {
		listenerStore.fire(new Event(Consts.ON_SHEET_CUSTOM_SORT_DIALOG));
	}
	public void openHyperlinkDialog() {
		listenerStore.fire(new Event(Consts.ON_SHEET_HYPERLINK_DIALOG));
	}
}