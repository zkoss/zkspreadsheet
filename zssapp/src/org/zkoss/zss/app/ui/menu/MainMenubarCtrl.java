/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ui.menu;

import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.select.Selectors;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.app.ui.CtrlBase;
import org.zkoss.zss.app.ui.AppEvts;
import org.zkoss.zss.app.ui.UiUtil;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.Version;
import org.zkoss.zss.ui.sys.UndoableActionManager;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menubar;
import org.zkoss.zul.Menuitem;
/**
 * 
 * @author dennis
 *
 */
public class MainMenubarCtrl extends CtrlBase<Menubar> {

	public MainMenubarCtrl() {
		super(true);
	}
	
	@Wire
	Menuitem newFile;
	@Wire
	Menuitem openManageFile;
	@Wire
	Menuitem saveFile;
	@Wire
	Menuitem saveFileAs;
	@Wire
	Menuitem saveFileAndClose;
	@Wire
	Menuitem closeFile;
	@Wire
	Menuitem exportFile;
	@Wire
	Menuitem exportPdf;
	
	
	@Wire
	Menuitem undo;
	@Wire
	Menuitem redo;
	
	
	@Wire
	Menuitem toggleFormulaBar;
	@Wire
	Menuitem freezePanel;
	@Wire
	Menuitem unfreezePanel;
	@Wire
	Menu freezeRows;
	@Wire
	Menu freezeCols;
	
	
	protected void onAppEvent(String event,Object data){
		if(AppEvts.ON_CHANGED_SPREADSHEET.equals(event)){
			doUpdateMenu((Spreadsheet)data);
		}else if(AppEvts.ON_UPDATE_UNDO_REDO.equals(event)){
			doUpdateMenu((Spreadsheet)data);
		}
		
	}
	
	private void doUpdateMenu(Spreadsheet sparedsheet){

		boolean hasBook = sparedsheet.getBook()!=null;
		boolean isEE = "EE".equals(Version.getEdition());
		//new and open are always on
		newFile.setDisabled(false);
		openManageFile.setDisabled(false);
		boolean readonly = UiUtil.isRepositoryReadonly();
		boolean disabled = !hasBook;
		saveFile.setDisabled(disabled || readonly);
		saveFileAs.setDisabled(disabled || readonly);
		saveFileAndClose.setDisabled(disabled || readonly);
		closeFile.setDisabled(disabled);
		exportFile.setDisabled(disabled);
		exportPdf.setDisabled(!isEE || disabled);
		
		
		
		UndoableActionManager uam = sparedsheet.getUndoableActionManager();
		
		
		String label = Labels.getLabel("zssapp.mainMenu.edit.undo");
		if(isEE && uam.isUndoable()){
			undo.setDisabled(false);
			label = label+":"+uam.getUndoLabel();	
		}else{
			undo.setDisabled(true);
		}
		undo.setLabel(label);
		
		label = Labels.getLabel("zssapp.mainMenu.edit.redo");
		if(isEE && uam.isRedoable()){
			redo.setDisabled(false);
			label = label+":"+uam.getRedoLabel();	
		}else{
			redo.setDisabled(true);
		}
		redo.setLabel(label);
				
//		toggleFormulaBar.setDisabled(disabled); //don't need to care the book load or not.
		toggleFormulaBar.setChecked(sparedsheet.isShowFormulabar());
		
		
		
		freezePanel.setDisabled(!isEE || disabled);
		Sheet sheet = sparedsheet.getSelectedSheet();
		unfreezePanel.setDisabled(!isEE || disabled || !(sheet.getRowFreeze()>0||sheet.getColumnFreeze()>0));
		
		for(Component comp:Selectors.find(freezeRows, "menuitem")){
			((Menuitem)comp).setDisabled(!isEE || disabled);
		}
		for(Component comp:Selectors.find(freezeCols, "menuitem")){
			((Menuitem)comp).setDisabled(!isEE || disabled);
		}
	}

	@Listen("onClick=#newFile")
	public void onNew(){
		pushAppEvent(AppEvts.ON_NEW_BOOK);
	}
	@Listen("onClick=#openManageFile")
	public void onOpen(){
		pushAppEvent(AppEvts.ON_OPEN_MANAGE_BOOK);
	}
	@Listen("onClick=#saveFile")
	public void onSave(){
		pushAppEvent(AppEvts.ON_SAVE_BOOK);
	}
	@Listen("onClick=#saveFileAs")
	public void onSaveAs(){
		pushAppEvent(AppEvts.ON_SAVE_BOOK_AS);
	}
	@Listen("onClick=#saveFileAndClose")
	public void onSaveClose(){
		pushAppEvent(AppEvts.ON_SAVE_CLOSE_BOOK);
	}
	@Listen("onClick=#closeFile")
	public void onClose(){
		pushAppEvent(AppEvts.ON_CLOSE_BOOK);
	}
	@Listen("onClick=#exportFile")
	public void onExport(){
		pushAppEvent(AppEvts.ON_EXPORT_BOOK);
	}
	@Listen("onClick=#exportPdf")
	public void onExportPdf(){
		pushAppEvent(AppEvts.ON_EXPORT_BOOK_PDF);
	}
	
	@Listen("onToggleFormulaBar=#mainMenubar")
	public void onToggleFormulaBar(){
		pushAppEvent(AppEvts.ON_TOGGLE_FORMULA_BAR);
	}
	
	@Listen("onFreezePanel=#mainMenubar")
	public void onFreezePanel(){
		pushAppEvent(AppEvts.ON_FREEZE_PNAEL);
	}
	
	@Listen("onUnfreezePanel=#mainMenubar")
	public void onUnfreezePanel(){
		pushAppEvent(AppEvts.ON_UNFREEZE_PANEL);
	}
	
	@Listen("onViewFreezeRows=#mainMenubar")
	public void onViewFreezeRows(ForwardEvent event) {
		int index = Integer.parseInt((String) event.getData());
		pushAppEvent(AppEvts.ON_FREEZE_ROW,index);
	}
	
	@Listen("onViewFreezeCols=#mainMenubar")
	public void onViewFreezeCols(ForwardEvent event) {
		int index = Integer.parseInt((String) event.getData());
		pushAppEvent(AppEvts.ON_FREEZE_COLUMN,index);
	}
	
	@Listen("onUndo=#mainMenubar")
	public void onUndo(ForwardEvent event) {
		pushAppEvent(AppEvts.ON_UNDO);
	}
	
	@Listen("onRedo=#mainMenubar")
	public void onRedo(ForwardEvent event) {
		pushAppEvent(AppEvts.ON_REDO);
	}
}
