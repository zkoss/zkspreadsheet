package org.zkoss.zss.app.ui;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.WebApps;
import org.zkoss.zk.ui.event.Event;
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
import org.zkoss.zss.app.ui.dlg.Dlgs;
import org.zkoss.zss.ui.DefaultUserAction;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zss.ui.UserActionHandler;
import org.zkoss.zssex.ui.DefaultExUserActionHandler;
import org.zkoss.zul.Messagebox;

public class AppCtrl extends CtrlBase<Component>{

	@Wire
	Spreadsheet ss;
	
	
	BookInfo selectedBookInfo;
	Book loadedBook;

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
				doSaveBook();
				return true;
			} 

			return super.dispatchAction(action);
		}
		
	}

	
	public void doLoadBook(){
		
	}
	
	
	public void doOpenNewBook(){
		selectedBookInfo = null;
		Importer importer = Importers.getImporter();
		
		try {
			loadedBook = importer.imports(WebApps.getCurrent().getResource("/WEB-INF/blank.xlsx"), "blank");
			MessageUtil.showInfoMessage("Loaded a new blank book");
		} catch (IOException e) {
			e.printStackTrace();
			MessageUtil.showWarnMessage("Can't load a new book");
			return;
		}
		
		ss.setBook(loadedBook);
		syncPageInfo();
	}
	
	private void syncPageInfo(){
		
		String name = selectedBookInfo!=null?selectedBookInfo.getName():loadedBook!=null?loadedBook.getBookName():null;
		getPage().setTitle(name!=null?"Book:"+name:"");
		
	}
	
	public void doOpenBook(){
		
		
		
		
		ss.setBook(loadedBook);
	}
	
	public void doSaveBook(){
		if(selectedBookInfo==null){
			doSaveBookAs();
			return;
		}
		BookRepository rep = getRepository();
		try {
			rep.save(selectedBookInfo, loadedBook);
			MessageUtil.showInfoMessage("Saved book "+selectedBookInfo.getName());
			syncPageInfo();
		} catch (IOException e) {
			e.printStackTrace();
			MessageUtil.showWarnMessage("Can't save book");
			return;
		}
	}


	
	public void doSaveBookAs(){
		String name = "New Book";
		if(selectedBookInfo!=null){
			name = "Copy Of "+selectedBookInfo.getName();
		}
		name = BookUtil.suggestBookName(name);
		
		Map<String,Object> args = newMap(newEntry("name",name),newEntry("callback",new EventListener<DlgCallbackEvent>(){
			public void onEvent(DlgCallbackEvent event) throws Exception {
				if(Dlgs.ON_SAVE.equals(event.getName())){
					String name = (String)event.getData("name");
					BookRepository rep = getRepository();
					try {
						selectedBookInfo = rep.save(name, loadedBook);
						MessageUtil.showInfoMessage("Saved book "+selectedBookInfo.getName());
						syncPageInfo();
					} catch (IOException e) {
						e.printStackTrace();
						MessageUtil.showWarnMessage("Can't save book");
						return;
					}
				}
			}}));
		Executions.createComponents("/zssapp/dlg/saveBookAs.zul", getSelf(), args);
	}
	
	
	
}
