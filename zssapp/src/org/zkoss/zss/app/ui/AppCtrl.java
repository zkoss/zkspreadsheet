/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ui;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.Map;
import java.util.Set;

import org.zkoss.lang.Classes;
import org.zkoss.lang.Strings;
import org.zkoss.util.logging.Log;
import org.zkoss.util.media.AMedia;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Execution;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.AreaRef;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.app.repository.BookInfo;
import org.zkoss.zss.app.repository.BookRepository;
import org.zkoss.zss.app.repository.BookRepositoryFactory;
import org.zkoss.zss.app.repository.impl.BookUtil;
import org.zkoss.zss.app.ui.dlg.DlgCallbackEvent;
import org.zkoss.zss.app.ui.dlg.OpenManageBookCtrl;
import org.zkoss.zss.app.ui.dlg.SaveBookAsCtrl;
import org.zkoss.zss.ui.AuxAction;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.UserActionContext;
import org.zkoss.zss.ui.UserActionHandler;
import org.zkoss.zss.ui.UserActionManager;
import org.zkoss.zss.ui.Version;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.impl.DefaultUserActionManagerCtrl;
import org.zkoss.zss.ui.sys.UndoableActionManager;
import org.zkoss.zss.ui.sys.UserActionManagerCtrl;
import org.zkoss.zss.ui.sys.SpreadsheetCtrl;
import org.zkoss.zul.Filedownload;

/**
 * 
 * @author dennis
 *
 */
public class AppCtrl extends CtrlBase<Component>{
	private static final Log log = Log.lookup(AppCtrl.class); 
	private static final long serialVersionUID = 1L;


	@Wire
	Spreadsheet ss;
	
	
	BookInfo selectedBookInfo;
	Book loadedBook;

	public AppCtrl() {
		super(true);
	}
	public void doBeforeComposeChildren(Component comp) throws Exception {
		super.doBeforeComposeChildren(comp);
		comp.setAttribute(APPCOMP, comp);
	}
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		boolean isEE = "EE".equals(Version.getEdition());
		boolean readonly = UiUtil.isRepositoryReadonly();
		UserActionManager uam = ss.getUserActionManager();
		uam.registerHandler(DefaultUserActionManagerCtrl.Category.AUXACTION.getName(), AuxAction.NEW_BOOK.getAction(), new UserActionHandler() {
			
			@Override
			public boolean process(UserActionContext ctx) {
				doOpenNewBook();
				return true;
			}
			
			@Override
			public boolean isEnabled(Book book, Sheet sheet) {
				return true;
			}
		});
		if(!readonly){
			uam.setHandler(DefaultUserActionManagerCtrl.Category.AUXACTION.getName(), AuxAction.SAVE_BOOK.getAction(), new UserActionHandler() {
				
				@Override
				public boolean process(UserActionContext ctx) {
					doSaveBook(false);
					return true;
				}
				
				@Override
				public boolean isEnabled(Book book, Sheet sheet) {
					return book!=null;
				}
			});
		}
		if(isEE){
			uam.setHandler(DefaultUserActionManagerCtrl.Category.AUXACTION.getName(), AuxAction.EXPORT_PDF.getAction(), new UserActionHandler() {
				
				@Override
				public boolean process(UserActionContext ctx) {
					doExportPdf();
					return true;
				}
				
				@Override
				public boolean isEnabled(Book book, Sheet sheet) {
					return book!=null;
				}
			});
		}
		
		//do after default
		uam.registerHandler(DefaultUserActionManagerCtrl.Category.AUXACTION.getName(), AuxAction.CLOSE_BOOK.getAction(), new UserActionHandler() {
			
			@Override
			public boolean process(UserActionContext ctx) {
				doCloseBook();
				return true;
			}
			
			@Override
			public boolean isEnabled(Book book, Sheet sheet) {
				return book!=null;
			}
		});
		
		ss.addEventListener(Events.ON_SHEET_SELECT, new EventListener<Event>() {
			@Override
			public void onEvent(Event event) throws Exception {
				onSheetSelect();
			}
		});
		
		ss.addEventListener(Events.ON_AFTER_UNDOABLE_MANAGER_ACTION, new EventListener<Event>() {
			@Override
			public void onEvent(Event event) throws Exception {
				onAfterUndoableManagerAction();
			}
		});
		
		
		//load default open book from parameter
		String bookName = null ;
		Execution exec = Executions.getCurrent();
		bookName = (String)exec.getArg().get("book");
		if(bookName==null){
			bookName = exec.getParameter("book");
		}
		
		BookInfo bookinfo = null;
		if(!Strings.isBlank(bookName)){
			bookName = bookName.trim();
			for(BookInfo info:getRepository().list()){
				if(bookName.equals(info.getName())){
					bookinfo = info;
					break;
				}
			}
		}
		String sheetName = Executions.getCurrent().getParameter("sheet");
		if(bookinfo!=null){
			doLoadBook(bookinfo,null,sheetName);
		}else{
			doOpenNewBook();
		}
	}
	
	
	private BookRepository getRepository(){
		return BookRepositoryFactory.getInstance().getRepository();
	}
	
	/*package*/ void doOpenNewBook(){
		selectedBookInfo = null;
		Importer importer = Importers.getImporter();
		
		try {
			loadedBook = importer.imports(getClass().getResourceAsStream("/web/zssapp/blank.xlsx"), "blank.xlsx");
		} catch (IOException e) {
			log.error(e.getMessage(),e);
			UiUtil.showWarnMessage("Can't load a new book");
			return;
		}
		ss.setBook(loadedBook);
		pushAppEvent(AppEvts.ON_LOADED_BOOK,loadedBook);
		pushAppEvent(AppEvts.ON_CHANGED_SPREADSHEET,ss);
		updatePageInfo();
//		UiUtil.showInfoMessage("Loaded a new blank book");
		
	}
	
	private void updatePageInfo(){
		
		String name = selectedBookInfo!=null?selectedBookInfo.getName():loadedBook!=null?loadedBook.getBookName():null;
		getPage().setTitle(name!=null?"Book : "+name:"");
		
	}
	
	/*package*/ void doCloseBook(){
		ss.setBook(loadedBook = null);
		selectedBookInfo = null;
		pushAppEvent(AppEvts.ON_CLOSED_BOOK,null);
		pushAppEvent(AppEvts.ON_CHANGED_SPREADSHEET,ss);
		updatePageInfo();
	}
	
	/*package*/ void onSheetSelect(){
		pushAppEvent(AppEvts.ON_CHANGED_SPREADSHEET,ss);
	}
	
	/*package*/ void onAfterUndoableManagerAction(){
		pushAppEvent(AppEvts.ON_UPDATE_UNDO_REDO,ss);
	}
	
	/*package*/ void doSaveBook(boolean close){
		if(UiUtil.isRepositoryReadonly()){
			return;
		}
		if(loadedBook==null){
			UiUtil.showWarnMessage("Please load a book first before save it");
			return;
		}
		if(selectedBookInfo==null){
			doSaveBookAs(close);
			return;
		}
		BookRepository rep = getRepository();
		try {
			rep.save(selectedBookInfo, loadedBook);
			UiUtil.showInfoMessage("Save book to "+selectedBookInfo.getName());
			pushAppEvent(AppEvts.ON_SAVED_BOOK,loadedBook);
			if(close){
				doCloseBook();
			}else{
				updatePageInfo();
			}
		} catch (IOException e) {
			log.error(e.getMessage(),e);
			UiUtil.showWarnMessage("Can't save book");
			return;
		}
	}


	
	private void doSaveBookAs(final boolean close){
		if(UiUtil.isRepositoryReadonly()){
			return;
		}
		if(loadedBook==null){
			UiUtil.showWarnMessage("Please load a book first before save it");
			return;
		}
		String name = "New Book";
		if(selectedBookInfo!=null){
			name = "Copy Of "+selectedBookInfo.getName();
		}
		name = BookUtil.suggestFileName(name,getRepository());
		
		SaveBookAsCtrl.show(new EventListener<DlgCallbackEvent>(){
			public void onEvent(DlgCallbackEvent event) throws Exception {
				if(SaveBookAsCtrl.ON_SAVE.equals(event.getName())){
					String name = (String)event.getData(SaveBookAsCtrl.ARG_NAME);
					BookRepository rep = getRepository();
					try {
						selectedBookInfo = rep.saveAs(name, loadedBook);
						pushAppEvent(AppEvts.ON_SAVED_BOOK,loadedBook);
						if(close){
							doCloseBook();
						}else{
							updatePageInfo();
						}
						UiUtil.showInfoMessage("Save book to "+selectedBookInfo.getName());
					} catch (IOException e) {
						log.error(e.getMessage(),e);
						UiUtil.showWarnMessage("Can't save book");
						return;
					}
				}
			}},name);
	}
	
	
	private void doLoadBook(BookInfo info,Book book,String sheetName){
		if(book==null){
			BookRepository rep = getRepository();
			try {
				book = rep.load(info);
			}catch (IOException e) {
				log.error(e.getMessage(),e);
				UiUtil.showWarnMessage("Can't load book");
				return;
			}
		}
		
		selectedBookInfo = info;
		loadedBook = book;
		
		ss.setBook(loadedBook);
		if(!Strings.isBlank(sheetName)){
			if(loadedBook.getSheet(sheetName)!=null){
				ss.setSelectedSheet(sheetName);
			}
		}
		
		pushAppEvent(AppEvts.ON_LOADED_BOOK,loadedBook);
		pushAppEvent(AppEvts.ON_CHANGED_SPREADSHEET,ss);
		updatePageInfo();
		
	}
	
	private void doOpenManageBook(){
		OpenManageBookCtrl.show(new EventListener<DlgCallbackEvent>(){
			public void onEvent(DlgCallbackEvent event) throws Exception {
				if(OpenManageBookCtrl.ON_OPEN.equals(event.getName())){
					BookInfo info = (BookInfo)event.getData(OpenManageBookCtrl.ARG_BOOKINFO);
					Book book = (Book)event.getData(OpenManageBookCtrl.ARG_BOOK);
					doLoadBook(info,book,null);
				}
			}});
	}
	
	private void doExportBook(){
		if(loadedBook==null){
			UiUtil.showWarnMessage("Please load a book first before export it");
			return;
		}
		String name = BookUtil.suggestFileName(loadedBook);
		File file;
		try {
			file = BookUtil.saveBookToWorkingFolder(loadedBook);
			Filedownload.save(new AMedia(name, null, "application/vnd.ms-excel", file, true));
		} catch (IOException e) {
			log.error(e.getMessage(),e);
			UiUtil.showWarnMessage("Can't export the book");
		}
	}
	
	/*package*/ void doExportPdf(){
		if(loadedBook==null){
			UiUtil.showWarnMessage("Please load a book first before export it");
			return;
		}
		String name = BookUtil.suggestPdfName(loadedBook);
		File file;
		try {
			file = BookUtil.saveBookToWorkingFolder(loadedBook,"pdf");
			Filedownload.save(new AMedia(name, null, "application/pdf", file, true));
		} catch (IOException e) {
			log.error(e.getMessage(),e);
			UiUtil.showWarnMessage("Can't export the book");
		}
	}
	
	@Override
	protected void onAppEvent(String event,Object data){
		//menu
		if(AppEvts.ON_NEW_BOOK.equals(event)){
			doOpenNewBook();
		}else if(AppEvts.ON_SAVE_BOOK.equals(event)){
			doSaveBook(false);
		}else if(AppEvts.ON_SAVE_BOOK_AS.equals(event)){
			doSaveBookAs(false);
		}else if(AppEvts.ON_SAVE_CLOSE_BOOK.equals(event)){
			doSaveBook(true);
		}else if(AppEvts.ON_CLOSE_BOOK.equals(event)){
			doCloseBook();
		}else if(AppEvts.ON_OPEN_MANAGE_BOOK.equals(event)){
			doOpenManageBook();
		}else if(AppEvts.ON_EXPORT_BOOK.equals(event)){
			doExportBook();
		}else if(AppEvts.ON_EXPORT_BOOK_PDF.equals(event)){
			doExportPdf();
		}else if(AppEvts.ON_TOGGLE_FORMULA_BAR.equals(event)){
			doToggleFormulabar();
		}else if(AppEvts.ON_FREEZE_PNAEL.equals(event)){
			AreaRef sel = ss.getSelection();
			doFreeze(sel.getRow(),sel.getColumn());
		}else if(AppEvts.ON_UNFREEZE_PANEL.equals(event)){
			doFreeze(0,0);
		}else if(AppEvts.ON_FREEZE_ROW.equals(event)){
			doFreeze(((Integer)data),ss.getSelectedSheet().getColumnFreeze());
		}else if(AppEvts.ON_FREEZE_COLUMN.equals(event)){
			doFreeze(ss.getSelectedSheet().getRowFreeze(),((Integer)data));
		}else if(AppEvts.ON_UNDO.equals(event)){
			doUndo();
		}else if(AppEvts.ON_REDO.equals(event)){
			doRedo();
		}
	}

	private void doUndo() {
		UndoableActionManager uam = ss.getUndoableActionManager();
		if(uam.isUndoable()){
			uam.undoAction();
		}
	}
	private void doRedo() {
		UndoableActionManager uam = ss.getUndoableActionManager();
		if(uam.isRedoable()){
			uam.redoAction();
		}
	} 
	
	private void doFreeze(int row, int column) {
		Ranges.range(ss.getSelectedSheet()).setFreezePanel(row, column);
		pushAppEvent(AppEvts.ON_CHANGED_SPREADSHEET,ss);
		
		//workaround before http://tracker.zkoss.org/browse/ZSS-390 fix
		AreaRef sel = ss.getSelection();
		row = row<0?sel.getRow():row;
		column = column<0?sel.getColumn():column;
		ss.setSelection(new AreaRef(row,column,row,column));
	}

	private void doToggleFormulabar() {
		ss.setShowFormulabar(!ss.isShowFormulabar());
		pushAppEvent(AppEvts.ON_CHANGED_SPREADSHEET,ss);
	}
	
	
}
