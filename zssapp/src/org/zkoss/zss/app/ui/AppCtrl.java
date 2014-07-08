/* 
	Purpose:
		
	Description:
		
	History:
		2013/7/10, Created by dennis

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.ui;

import java.io.*;

import org.zkoss.image.AImage;
import org.zkoss.lang.Strings;
import org.zkoss.util.logging.Log;
import org.zkoss.util.media.AMedia;
import org.zkoss.util.media.Media;
import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.*;
import org.zkoss.zk.ui.event.*;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.app.repository.*;
import org.zkoss.zss.app.repository.impl.BookUtil;
import org.zkoss.zss.app.ui.dlg.*;
import org.zkoss.zss.ui.*;
import org.zkoss.zss.ui.event.Events;
import org.zkoss.zss.ui.impl.DefaultUserActionManagerCtrl;
import org.zkoss.zss.ui.sys.UndoableActionManager;
import org.zkoss.zul.Filedownload;
import org.zkoss.zul.Fileupload;

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
						UiUtil.showInfoMessage("Save book to "+selectedBookInfo.getName());
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
			//default Excel exporter exports XLSX 
			Filedownload.save(new AMedia(name, null, "application/vnd.ms-excel.12", file, true));
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
		} else if (AppEvts.ON_INSERT_PICTURE.equals(event)) {
			doInsertPicture();
		} else if (AppEvts.ON_INSERT_CHART.equals(event)) {
			doInsertChart((String) data);
		} else if (AppEvts.ON_INSERT_HYPERLINK.equals(event)) {
			doInsertHyperlink();
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
	
	private void doInsertPicture() {
		Fileupload.get(1,new EventListener<UploadEvent>() {
			public void onEvent(UploadEvent event) throws Exception {
				Media media = event.getMedia();
				if(media == null){
					return;
				}
				if(!(media instanceof AImage) || SheetOperationUtil.getPictureFormat((AImage)media)==null){
					UiUtil.showWarnMessage(Labels.getLabel("zss.actionhandler.msg.cant_support_file", new Object[]{media==null?"":media.getName()}));
					return;
				}
				final Sheet sheet = ss.getSelectedSheet();
				final AreaRef selection = ss.getSelection();
				Range range = Ranges.range(sheet, selection.getRow(), selection.getColumn(), selection.getLastRow(), selection.getLastColumn());
				
				SheetOperationUtil.addPicture(range,(AImage)media);
				
				return;
			}
		});
	}
	
	private void doInsertChart(String type) {
		AreaRef selection = ss.getSelection();
		Range range = Ranges.range(ss.getSelectedSheet(), selection.getRow(),
				selection.getColumn(), selection.getLastRow(),
				selection.getLastColumn());
		SheetAnchor anchor = SheetOperationUtil.toChartAnchor(range);
		SheetOperationUtil.addChart(range,anchor, toChartType(type), Chart.Grouping.STANDARD, Chart.LegendPosition.RIGHT);
	}
	
	private Chart.Type toChartType(String type) {
		Chart.Type chartType;
		if ("ColumnChart".equals(type)) {
			chartType = Chart.Type.COLUMN;
		} else if ("ColumnChart3D".equals(type)) {
			chartType = Chart.Type.COLUMN_3D;
		} else if ("LineChart".equals(type)) {
			chartType = Chart.Type.LINE;
		} else if ("LineChart3D".equals(type)) {
			chartType = Chart.Type.LINE_3D;
		} else if ("PieChart".equals(type)) {
			chartType = Chart.Type.PIE;
		} else if ("PieChart3D".equals(type)) {
			chartType = Chart.Type.PIE_3D;
		} else if ("BarChart".equals(type)) {
			chartType = Chart.Type.BAR;
		} else if ("BarChart3D".equals(type)) {
			chartType = Chart.Type.BAR_3D;
		} else if ("AreaChart".equals(type)) {
			chartType = Chart.Type.AREA;
		} else if ("ScatterChart".equals(type)) {
			chartType = Chart.Type.SCATTER;
		} else if ("DoughnutChart".equals(type)) {
			chartType = Chart.Type.DOUGHNUT;
		} else {
			chartType = Chart.Type.LINE;
		}
		return chartType;
	}
	
	private void doInsertHyperlink() {
		final Sheet sheet = ss.getSelectedSheet();
		final AreaRef selection = ss.getSelection();
		final Range range = Ranges.range(sheet, selection);
		Hyperlink link = range.getCellHyperlink();
		String display = link == null ? range.getCellFormatText():link.getLabel();
		String address = link == null ? null:link.getAddress();
		HyperlinkCtrl.show(new EventListener<DlgCallbackEvent>(){
			public void onEvent(DlgCallbackEvent event) throws Exception {
				if(HyperlinkCtrl.ON_OK.equals(event.getName())){
					final String address = (String) event.getData(HyperlinkCtrl.ARG_ADDRESS);
					final String label = (String) event.getData(HyperlinkCtrl.ARG_DISPLAY);
					CellOperationUtil.applyHyperlink(range, HyperlinkType.URL, address, label);
				}
			}}, address, display);
	}
	
}
