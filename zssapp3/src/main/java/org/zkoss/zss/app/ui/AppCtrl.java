package org.zkoss.zss.app.ui;

import java.io.IOException;
import java.util.Map;
import java.util.Set;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zss.api.Importer;
import org.zkoss.zss.api.Importers;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.app.repository.BookInfo;
import org.zkoss.zss.app.repository.BookRepository;
import org.zkoss.zss.app.repository.BookRepositoryFactory;
import org.zkoss.zss.app.repository.BookUtil;
import org.zkoss.zss.app.ui.dlg.DlgCallbackEvent;
import org.zkoss.zss.app.ui.dlg.DlgEvts;
import org.zkoss.zss.ui.DefaultUserAction;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zssex.ui.DefaultExUserActionHandler;

public class AppCtrl extends CtrlBase<Component>{

	private static final long serialVersionUID = 1L;


	@Wire
	Spreadsheet ss;
	
	
	BookInfo selectedBookInfo;
	Book loadedBook;

	public AppCtrl() {
		super(true);
	}
	
	@Override
	public void doAfterCompose(Component comp) throws Exception {
		super.doAfterCompose(comp);
		
		ss.setUserActionHandler(new AppUserActionHandler());
		
		//TODO load default open book from parameter
		
		doOpenNewBook();
	}
	
	
	private BookRepository getRepository(){
		return BookRepositoryFactory.getInstance().getRepository();
	}
	
	
	public class AppUserActionHandler extends DefaultExUserActionHandler  {
		private static final long serialVersionUID = 1L;
		
		@Override
		public Set<String> getSupportedUserAction(Sheet sheet) {
			Set<String> actions = super.getSupportedUserAction(sheet);
			actions.add(DefaultUserAction.NEW_BOOK.getAction());
			if(sheet!=null){
				actions.add(DefaultUserAction.SAVE_BOOK.getAction());
			}
			return actions;
		}
		
		@Override
		protected boolean dispatchAction(String action) {
			DefaultUserAction dua = DefaultUserAction.getBy(action);

			if (DefaultUserAction.NEW_BOOK.equals(dua)) {
				doOpenNewBook();
				return true;
			} else if (DefaultUserAction.SAVE_BOOK.equals(dua)) {
				doSaveBook(false);
				return true;
			} 

			return super.dispatchAction(action);
		}
		
		@Override
		protected boolean doCloseBook(){
			super.doCloseBook();
			AppCtrl.this.doCloseBook();
			return true;
		}
	}
	
	
	private void doOpenNewBook(){
		selectedBookInfo = null;
		Importer importer = Importers.getImporter();
		
		try {
			loadedBook = importer.imports(WebApps.getCurrent().getResource("/WEB-INF/blank.xlsx"), "blank");
			UiUtil.showInfoMessage("Loaded a new blank book");
		} catch (IOException e) {
			e.printStackTrace();
			UiUtil.showWarnMessage("Can't load a new book");
			return;
		}
		ss.setBook(loadedBook);
		pushDesktopEvent(DesktopEvts.ON_LOADED_BOOK,loadedBook);
		updatePageInfo();
		
	}
	
	private void updatePageInfo(){
		
		String name = selectedBookInfo!=null?selectedBookInfo.getName():loadedBook!=null?loadedBook.getBookName():null;
		getPage().setTitle(name!=null?"Book : "+name:"");
		
	}
	
	private void doCloseBook(){
		ss.setBook(loadedBook = null);
		selectedBookInfo = null;
		pushDesktopEvent(DesktopEvts.ON_CLOSED_BOOK,null);
		updatePageInfo();
	}
	
	private void doSaveBook(boolean close){
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
			UiUtil.showInfoMessage("Saved book "+selectedBookInfo.getName());
			pushDesktopEvent(DesktopEvts.ON_SAVED_BOOK,loadedBook);
			if(close){
				doCloseBook();
			}else{
				updatePageInfo();
			}
		} catch (IOException e) {
			e.printStackTrace();
			UiUtil.showWarnMessage("Can't save book");
			return;
		}
	}


	
	private void doSaveBookAs(final boolean close){
		if(loadedBook==null){
			UiUtil.showWarnMessage("Please load a book first before save it");
			return;
		}
		String name = "New Book";
		if(selectedBookInfo!=null){
			name = "Copy Of "+selectedBookInfo.getName();
		}
		name = BookUtil.suggestBookName(name);
		
		Map<String,Object> args = newMap(newEntry("name",name),newEntry("callback",new EventListener<DlgCallbackEvent>(){
			public void onEvent(DlgCallbackEvent event) throws Exception {
				if(DlgEvts.ON_SAVE.equals(event.getName())){
					String name = (String)event.getData("name");
					BookRepository rep = getRepository();
					try {
						selectedBookInfo = rep.saveAs(name, loadedBook);
						UiUtil.showInfoMessage("Saved book "+selectedBookInfo.getName());
						pushDesktopEvent(DesktopEvts.ON_SAVED_BOOK,loadedBook);
						if(close){
							doCloseBook();
						}else{
							updatePageInfo();
						}
					} catch (IOException e) {
						e.printStackTrace();
						UiUtil.showWarnMessage("Can't save book");
						return;
					}
				}
			}}));
		Executions.createComponents("/zssapp/dlg/saveBookAs.zul", getSelf(), args);
	}
	
	
	private void doOpenManageBook(){
		Map<String,Object> args = newMap(newEntry("callback",new EventListener<DlgCallbackEvent>(){
			public void onEvent(DlgCallbackEvent event) throws Exception {
				if(DlgEvts.ON_OPEN.equals(event.getName())){
					BookInfo info = (BookInfo)event.getData("bookinfo");
					Book book = (Book)event.getData("book");
					if(book==null){
						BookRepository rep = getRepository();
						try {
							book = rep.load(info);
						}catch (IOException e) {
							e.printStackTrace();
							UiUtil.showWarnMessage("Can't load book");
							return;
						}
					}
					
					selectedBookInfo = info;
					loadedBook = book;
					
					ss.setBook(loadedBook);
					pushDesktopEvent(DesktopEvts.ON_LOADED_BOOK,loadedBook);
				}
			}}));
		Executions.createComponents("/zssapp/dlg/openBook.zul", getSelf(), args);
	}
	
	@Override
	protected void onDesktopEvent(String event,Object data){
		System.out.println(">>>>>onDesktop "+event+","+data);
		
		//menu
		if(DesktopEvts.ON_NEW_BOOK.equals(event)){
			doOpenNewBook();
		}else if(DesktopEvts.ON_SAVE_BOOK.equals(event)){
			doSaveBook(false);
		}else if(DesktopEvts.ON_SAVE_BOOK_AS.equals(event)){
			doSaveBookAs(false);
		}else if(DesktopEvts.ON_SAVE_CLOSE_BOOK.equals(event)){
			doSaveBook(true);
		}else if(DesktopEvts.ON_CLOSE_BOOK.equals(event)){
			doCloseBook();
		}else if(DesktopEvts.ON_OPEN_MANAGE_BOOK.equals(event)){
			doOpenManageBook();
		}else if(DesktopEvts.ON_EXPORT_BOOK.equals(event)){
			
		}
	}
	
	
}
