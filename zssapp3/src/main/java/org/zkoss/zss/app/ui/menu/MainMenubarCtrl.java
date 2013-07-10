package org.zkoss.zss.app.ui.menu;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.Selectors;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.app.ui.CtrlBase;
import org.zkoss.zss.app.ui.DesktopEvts;
import org.zkoss.zul.Menu;
import org.zkoss.zul.Menubar;
import org.zkoss.zul.Menuitem;

public class MainMenubarCtrl extends CtrlBase<Menubar> {

	public MainMenubarCtrl() {
		super(true);
	}
	
	@Wire
	Menuitem newFile;
	@Wire
	Menuitem openFile;
	@Wire
	Menuitem saveFile;
	@Wire
	Menuitem saveFileAs;
	@Wire
	Menuitem saveFileAndClose;
	@Wire
	Menuitem closeFile;
	@Wire
	Menuitem deleteFile;
	@Wire
	Menuitem importFile;
	@Wire
	Menuitem exportFile;
	
	
	@Wire
	Menuitem viewFormulaBar;
	@Wire
	Menu freezeRows;
	@Wire
	Menu freezeCols;
	
	
	
	protected void onDesktopEvent(String event,Object data){
		if(DesktopEvts.ON_CLOSED_BOOK.equals(event)){
			doUpdateMenu(false);
		}else if(DesktopEvts.ON_LOADED_BOOK.equals(event)){
			doUpdateMenu(true);
		}
	}
	
	private void doUpdateMenu(boolean hasBook) {
		//new and import are always on
		newFile.setDisabled(false);
		importFile.setDisabled(false);
		
		boolean disabled = !hasBook;
		openFile.setDisabled(disabled);
		saveFile.setDisabled(disabled);
		saveFileAs.setDisabled(disabled);
		saveFileAndClose.setDisabled(disabled);
		closeFile.setDisabled(disabled);
		deleteFile.setDisabled(disabled);
		exportFile.setDisabled(disabled);
		
		
		viewFormulaBar.setDisabled(disabled);
		for(Component comp:Selectors.find(freezeRows, "menuitem")){
			((Menuitem)comp).setDisabled(disabled);
		}
		for(Component comp:Selectors.find(freezeCols, "menuitem")){
			((Menuitem)comp).setDisabled(disabled);
		}
	}


	@Listen("onClick=#newFile")
	public void onNew(){
		pushDesktopEvent(DesktopEvts.ON_NEW_BOOK, null);
	}
	@Listen("onClick=#saveFile")
	public void onSave(){
		pushDesktopEvent(DesktopEvts.ON_SAVE_BOOK, null);
	}
	@Listen("onClick=#saveFileAs")
	public void onSaveAs(){
		pushDesktopEvent(DesktopEvts.ON_SAVE_BOOK_AS, null);
	}
	@Listen("onClick=#saveFileAndClose")
	public void onSaveClose(){
		pushDesktopEvent(DesktopEvts.ON_SAVE_CLOSE_BOOK, null);
	}
	@Listen("onClick=#closeFile")
	public void onClose(){
		pushDesktopEvent(DesktopEvts.ON_CLOSE_BOOK, null);
	}
	@Listen("onClick=#deleteFile")
	public void onDelete(){
		pushDesktopEvent(DesktopEvts.ON_DELETE_BOOK, null);
	}
	@Listen("onClick=#importFile")
	public void onImport(){
		pushDesktopEvent(DesktopEvts.ON_IMPORT_BOOK, null);
	}
	@Listen("onClick=#exportFile")
	public void onExport(){
		pushDesktopEvent(DesktopEvts.ON_EXPORT_BOOK, null);
	}
}
